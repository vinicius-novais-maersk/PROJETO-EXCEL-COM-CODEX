# Excel + OpenAI Local Integration (FastAPI + VBA)

Este projeto cria uma integração local entre Excel e OpenAI sem usar suplemento da Microsoft.

O fluxo funciona assim:

1. O Excel executa uma macro VBA.
2. A macro lê o texto da célula ativa.
3. A macro envia esse texto para uma API HTTP local em `http://127.0.0.1:8000/ask`.
4. O backend em Python recebe o texto.
5. O backend chama a OpenAI Responses API usando o SDK oficial.
6. O backend devolve a resposta em JSON.
7. A macro escreve a resposta na célula ao lado.

## Estrutura do projeto

```text
excel-openai-local/
|-- .env.example
|-- .gitignore
|-- app.py
|-- config.py
|-- requirements.txt
|-- README.md
`-- vba/
    `-- ExcelGPT.bas
```

## O que cada arquivo faz

- `app.py`: backend FastAPI com os endpoints HTTP e chamada para a OpenAI.
- `config.py`: leitura das variáveis de ambiente e configuração básica.
- `.env.example`: modelo do arquivo `.env` que você deve criar localmente.
- `requirements.txt`: dependências Python do projeto.
- `vba/ExcelGPT.bas`: módulo VBA pronto para importar no Excel.
- `README.md`: documentação completa para instalar, executar e testar no Windows.

## Requisitos

- Windows com Excel Desktop
- Python 3.10 ou superior
- VSCode
- Uma chave de API da OpenAI configurada em variável de ambiente

## Endpoint principal

### `POST /ask`

Exemplo de entrada:

```json
{
  "prompt": "texto enviado pelo Excel",
  "system_prompt": "opcional"
}
```

Exemplo de sucesso:

```json
{
  "success": true,
  "answer": "texto da resposta"
}
```

Exemplo de erro:

```json
{
  "success": false,
  "error": "mensagem"
}
```

## Passo a passo no Windows

### 1. Abrir o projeto no VSCode

Abra a pasta do projeto no VSCode.

### 2. Criar o ambiente virtual

No terminal do VSCode, execute:

```powershell
py -3.11 -m venv .venv
```

Se o seu Python estiver associado a outro comando, você pode usar:

```powershell
python -m venv .venv
```

### 3. Ativar o ambiente virtual

No PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Se o PowerShell bloquear a ativação por política de execução, rode esta linha apenas na sessão atual:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

Depois execute novamente:

```powershell
.venv\Scripts\Activate.ps1
```

### 4. Instalar as dependências

Com o ambiente virtual ativo:

```powershell
pip install -r requirements.txt
```

### 5. Configurar a API key

Crie o arquivo `.env` a partir do exemplo:

```powershell
Copy-Item .env.example .env
```

Abra o arquivo `.env` e preencha sua chave:

```env
OPENAI_API_KEY=your_openai_api_key_here
APP_HOST=127.0.0.1
APP_PORT=8000
OPENAI_MODEL=gpt-5.4
OPENAI_TIMEOUT_SECONDS=60
LOG_LEVEL=INFO
```

Troque apenas o valor de `OPENAI_API_KEY` pela sua chave real.

Importante:

- Nunca coloque a chave direto no código.
- O backend já lê a chave do ambiente usando `.env`.

### 6. Rodar o backend local

Ainda no terminal com o ambiente virtual ativo:

```powershell
python app.py
```

Se tudo estiver certo, a API ficará disponível em:

```text
http://127.0.0.1:8000
```

A documentação automática do FastAPI ficará em:

```text
http://127.0.0.1:8000/docs
```

### 7. Testar o backend antes do Excel

Você pode testar com PowerShell:

```powershell
$body = @{
    prompt = "Write a short greeting for Excel."
    system_prompt = "Answer in Brazilian Portuguese."
} | ConvertTo-Json

Invoke-RestMethod `
    -Uri "http://127.0.0.1:8000/ask" `
    -Method Post `
    -ContentType "application/json" `
    -Body $body
```

Se estiver tudo certo, você receberá um JSON com `success = true` e a resposta em `answer`.

## Como importar o módulo VBA no Excel

### 1. Salvar a planilha como `.xlsm`

Para usar macros, a planilha precisa ser salva em formato habilitado para macro:

```text
Excel Macro-Enabled Workbook (*.xlsm)
```

### 2. Abrir o Editor VBA

No Excel:

- pressione `Alt + F11`

### 3. Importar o módulo

No Editor VBA:

1. Vá em `File > Import File...`
2. Selecione o arquivo `vba\ExcelGPT.bas`

O módulo será importado com o nome `ExcelGPT`.

### 4. Habilitar macros

Ao abrir a planilha, clique em `Enable Content` se o Excel pedir confirmação.

## Como usar a macro `PerguntarGPT`

1. Escreva um prompt em uma célula, por exemplo `A1`.
2. Selecione essa célula.
3. Execute a macro `PerguntarGPT`.
4. A macro enviará o conteúdo da célula para a API local.
5. A resposta será escrita na célula ao lado, por exemplo `B1`.

## Como executar a macro

Você pode rodar a macro de uma destas formas:

- No Excel, pressione `Alt + F8`, selecione `PerguntarGPT` e clique em `Run`.
- Ou crie um botão na planilha e associe a macro `PerguntarGPT`.

## Sobre o JSON no VBA

Neste projeto, o módulo VBA já inclui um parser mínimo para o formato de resposta usado por esta API.

Isso significa que:

- você não precisa instalar biblioteca extra de JSON para o cenário básico;
- a macro já consegue ler os campos `success`, `answer` e `error`;
- isso mantém a solução mais simples para Excel Desktop no Windows.

Se no futuro você quiser trabalhar com JSON genérico e estruturas mais complexas, uma opção comum é usar a biblioteca `VBA-JSON`. Mas para esta integração específica, ela não é necessária.

## Tratamento de erros incluído

O projeto já trata:

- prompt vazio;
- chave `OPENAI_API_KEY` ausente;
- erro de autenticação com a OpenAI;
- rate limit;
- timeout;
- falha de conexão com a OpenAI;
- backend local fora do ar;
- resposta vazia;
- erro amigável no Excel quando a célula ativa estiver vazia.

## Logs

O backend usa logs básicos no terminal para facilitar depuração:

- recebimento de requisição;
- tamanho do prompt;
- uso ou não de `system_prompt`;
- sucesso na geração da resposta;
- falhas de autenticação, timeout, conexão e outros erros.

## Arquivos principais

### `app.py`

Responsável por:

- criar a aplicação FastAPI;
- expor `POST /ask`;
- validar o JSON recebido;
- chamar a OpenAI Responses API com `model="gpt-5.4"`;
- devolver JSON padronizado para o Excel.

### `config.py`

Responsável por:

- carregar `.env`;
- centralizar host, porta, modelo e timeout;
- manter a configuração simples e organizada.

### `vba/ExcelGPT.bas`

Responsável por:

- ler o valor da célula ativa;
- enviar o prompt para o backend local;
- processar a resposta JSON;
- escrever a saída na célula ao lado;
- mostrar mensagens amigáveis em caso de erro.

## Teste completo da integração

### Teste 1: backend

1. Ative o ambiente virtual.
2. Rode `python app.py`.
3. Confirme que `http://127.0.0.1:8000/docs` abre normalmente.

### Teste 2: Excel

1. Abra a planilha `.xlsm`.
2. Digite em `A1`:

```text
Explique em uma frase o que é tabela dinâmica.
```

3. Selecione `A1`.
4. Rode a macro `PerguntarGPT`.
5. Verifique se a resposta aparece em `B1`.

### Teste 3: erro amigável

1. Pare o backend Python.
2. Volte para o Excel.
3. Rode a macro novamente.
4. Você deverá ver uma mensagem amigável informando para verificar se o backend local está rodando.

## Observações finais

- O backend é local e simples, ideal para uso no VSCode e testes no próprio computador.
- O Excel fala com `127.0.0.1`, então o fluxo depende de o backend estar rodando na sua máquina.
- O código foi mantido enxuto e organizado para facilitar manutenção futura.
