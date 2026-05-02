import logging
from typing import Any, Dict, Optional

import uvicorn
from fastapi import FastAPI
from fastapi.responses import JSONResponse
from openai import (
    APIConnectionError,
    APITimeoutError,
    AuthenticationError,
    BadRequestError,
    OpenAI,
    OpenAIError,
    RateLimitError,
)
from pydantic import BaseModel, Field

from config import get_settings


settings = get_settings()

logging.basicConfig(
    level=getattr(logging, settings.log_level, logging.INFO),
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("excel_openai_api")

if not settings.openai_api_key:
    logger.warning(
        "OPENAI_API_KEY is not configured. The /ask endpoint will return an error until it is set."
    )


app = FastAPI(
    title="Excel OpenAI Local API",
    version="1.0.0",
    description="Local HTTP API that allows Excel VBA to send prompts to the OpenAI Responses API.",
)


class AskRequest(BaseModel):
    prompt: str = Field(..., description="Text received from Excel.")
    system_prompt: Optional[str] = Field(
        default=None,
        description="Optional system instruction sent to the model.",
    )


class AskSuccessResponse(BaseModel):
    success: bool = True
    answer: str


class AskErrorResponse(BaseModel):
    success: bool = False
    error: str


def create_openai_client() -> OpenAI:
    return OpenAI(
        api_key=settings.openai_api_key,
        timeout=settings.openai_timeout_seconds,
    )


def build_error_response(message: str, status_code: int) -> JSONResponse:
    return JSONResponse(
        status_code=status_code,
        content=AskErrorResponse(error=message).model_dump(),
    )


def extract_text_from_response(response: Any) -> str:
    output_text = getattr(response, "output_text", None)
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()

    collected_chunks = []
    for item in getattr(response, "output", []) or []:
        if getattr(item, "type", None) != "message":
            continue

        for content_item in getattr(item, "content", []) or []:
            if getattr(content_item, "type", None) == "output_text":
                text = getattr(content_item, "text", "")
                if text:
                    collected_chunks.append(text.strip())

    return "\n".join(chunk for chunk in collected_chunks if chunk).strip()


@app.get("/")
def root() -> Dict[str, str]:
    return {
        "service": "Excel OpenAI Local API",
        "status": "running",
        "docs": "/docs",
    }


@app.get("/health")
def health() -> Dict[str, bool]:
    return {"success": True}


@app.post(
    "/ask",
    response_model=AskSuccessResponse,
    responses={
        400: {"model": AskErrorResponse},
        401: {"model": AskErrorResponse},
        429: {"model": AskErrorResponse},
        500: {"model": AskErrorResponse},
        502: {"model": AskErrorResponse},
        504: {"model": AskErrorResponse},
    },
)
def ask(payload: AskRequest):
    prompt = payload.prompt.strip()
    system_prompt = (payload.system_prompt or "").strip()

    if not prompt:
        logger.warning("Rejected /ask request because prompt was empty.")
        return build_error_response("The 'prompt' field cannot be empty.", 400)

    if not settings.openai_api_key:
        logger.error("Rejected /ask request because OPENAI_API_KEY is missing.")
        return build_error_response(
            "OPENAI_API_KEY is not configured. Create a .env file and set your API key.",
            500,
        )

    logger.info(
        "Processing /ask request. Prompt length=%s, has_system_prompt=%s",
        len(prompt),
        bool(system_prompt),
    )

    request_params: Dict[str, Any] = {
        "model": settings.openai_model,
        "input": prompt,
    }

    if system_prompt:
        request_params["instructions"] = system_prompt

    try:
        response = create_openai_client().responses.create(**request_params)
        answer = extract_text_from_response(response)

        if not answer:
            logger.warning("OpenAI returned an empty text response.")
            return build_error_response("The model returned an empty response.", 502)

        logger.info("OpenAI response generated successfully. Answer length=%s", len(answer))
        return AskSuccessResponse(answer=answer)

    except AuthenticationError:
        logger.exception("OpenAI authentication failed.")
        return build_error_response(
            "Authentication failed. Check whether OPENAI_API_KEY is valid.",
            401,
        )
    except RateLimitError:
        logger.exception("OpenAI rate limit reached.")
        return build_error_response(
            "Rate limit reached. Wait a moment and try again.",
            429,
        )
    except BadRequestError as exc:
        logger.exception("OpenAI rejected the request.")
        return build_error_response(
            f"OpenAI rejected the request: {exc}",
            400,
        )
    except APITimeoutError:
        logger.exception("OpenAI request timed out.")
        return build_error_response(
            "The request to OpenAI timed out. Try again in a moment.",
            504,
        )
    except APIConnectionError:
        logger.exception("Could not connect to OpenAI.")
        return build_error_response(
            "Could not connect to OpenAI. Check your internet connection and try again.",
            502,
        )
    except OpenAIError as exc:
        logger.exception("Unexpected OpenAI SDK error.")
        return build_error_response(
            f"Unexpected OpenAI error: {exc}",
            502,
        )
    except Exception:
        logger.exception("Unexpected server error while processing /ask.")
        return build_error_response(
            "Unexpected internal server error.",
            500,
        )


if __name__ == "__main__":
    uvicorn.run(
        "app:app",
        host=settings.app_host,
        port=settings.app_port,
        reload=False,
        log_level=settings.log_level.lower(),
    )

