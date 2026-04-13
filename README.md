# 📊 Automação de Extração de Extrato Bancário (Python)

Este projeto realiza a automação da extração de dados de extrato
bancário via API, processando múltiplas contas a partir de um arquivo
Excel e gerando relatórios estruturados de saída.

A solução contempla autenticação, paginação de dados, tratamento de
erros e geração de logs, sendo ideal para cenários de automação
financeira e RPA.

------------------------------------------------------------------------

## 🚀 Principais Funcionalidades

-   🔐 Autenticação automática via API (OAuth / Client Credentials)
-   📥 Leitura de contas a partir de arquivo Excel
-   🔄 Paginação automática na coleta de dados
-   📊 Transformação dos dados em DataFrame (pandas)
-   📁 Geração de relatório Excel por conta
-   🧾 Registro de logs de execução (sucesso e erro)
-   ⚙️ Suporte a múltiplos ambientes (HML / PRD)

------------------------------------------------------------------------

## 🧱 Estrutura do Projeto

    .
    ├── bot.py              # Script principal (orquestração)
    ├── CustomLib.py        # Funções auxiliares e regras de negócio
    ├── Config/
    │   └── Config.env      # Configurações do processo
    ├── Input/
    │   └── input.xlsx      # Arquivo de entrada
    ├── Output/
    │   └── output.xlsx     # Arquivo gerado

------------------------------------------------------------------------

## ⚙️ Configuração

O projeto utiliza um arquivo `.env` para centralizar as configurações.

------------------------------------------------------------------------

## 📥 Formato do Arquivo de Entrada

O Excel de entrada deve conter as seguintes colunas:

|  Campo       | Descrição
| ------------ | -----------------------------
|  Agencia     | Número da agência
|  Conta       | Número da conta
|  DataInicio  | Data inicial (opcional)
|  DataFim     | Data final (opcional)
|  HomolId     | Obrigatório apenas para HML

------------------------------------------------------------------------

## ▶️ Execução

### 1. Instalar dependências

    pip install -r requirements.txt

### 2. Configurar o arquivo `.env`

    Preencha as credenciais da API.

### 3. Executar o processo

    python bot.py

------------------------------------------------------------------------

## 🧠 Fluxo do Processo

1.  Leitura do arquivo de configuração (`.env`)
2.  Leitura do Excel de entrada
3.  Para cada conta:
    -   Ajuste dos dados (agência, conta, datas)
    -   Geração/validação do token de acesso
    -   Consumo da API de extrato (com paginação)
    -   Conversão dos dados para DataFrame
    -   Escrita no Excel de saída
4.  Geração de aba de logs com status por conta

------------------------------------------------------------------------

## 📊 Saída Gerada

O arquivo Excel final contém:

-   📄 Uma aba por conta (`Agencia-Conta`)
-   🧾 Aba de **Logs** com status da execução

------------------------------------------------------------------------

## ⚠️ Tratamento de Erros

-   Falhas de API são capturadas e registradas
-   O processamento continua mesmo em caso de erro em uma conta
-   Mensagens de erro são armazenadas na aba de logs

------------------------------------------------------------------------

## 🛠️ Tecnologias Utilizadas

-   Python 3.x\
-   pandas\
-   requests\
-   python-dotenv\
-   openpyxl

------------------------------------------------------------------------

## 👤 Autor

Pedro Coletti
