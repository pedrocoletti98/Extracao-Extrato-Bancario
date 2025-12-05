import os
import json
import base64
import requests
import pandas as pd
import datetime as dt
from dotenv import dotenv_values

# Read Config Files and returns the dictionary used throughout the process.
# Input: Config File Path
def read_config(in_strConfigFilePath):
    # Check if Config File Exists
    if not os.path.exists(in_strConfigFilePath):
        raise Exception("Config File Not Found")
    # Load Config values
    return dotenv_values(in_strConfigFilePath)


# Read the Input Excel File based on Config settings and returns the dataframe of that Excel.
# Input: Config Dictionary
def read_excel(in_dictConfig):
    # Initialize and sets Excel variables
    strExcelFileName = in_dictConfig["InputExcelFileName"]
    strInputFolder = in_dictConfig["InputFolder"]
    strExcelFilePath = os.path.join(os.curdir, strInputFolder, strExcelFileName)
    strExcelSheetName = in_dictConfig["InputExcelSheetName"]

    # Check if Input Excel File Exists
    if not os.path.exists(strExcelFilePath):
        raise Exception("Excel File Not Found")

    # Read and returns excel dataframe
    return pd.read_excel(strExcelFilePath, strExcelSheetName)

# Set and adjusts input excel variables to be used on the request_extract_info function.
# Input: Config Dictionary and Row Item from the Input Excel
# Returns strAgencia, strConta, strDataInicio, strDataFim, strHomolId
def set_excel_variables(in_dictConfig, in_rowItem):
    # Get Environment Type from Config File (HML or PRD)
    strEnvType = in_dictConfig["EnvType"]

    # Get Current Excel Row data
    strAgencia = str(int(in_rowItem["Agencia"]))
    strConta = str(int(in_rowItem["Conta"]))
    strDataInicio = str(in_rowItem["DataInicio"])
    strDataFim = str(in_rowItem["DataFim"])

    # Check if both dates are None and adjust to correct format. Expected Input Format: "dd-mm-YYYY"
    # Check if dataIncio = None and adjusts correct type
    if strDataInicio == "nan":
        strDataInicio = None
    else:
        # Adjusts to ddmmYYYY
        "".join(strDataInicio.split("-")).lstrip("0")

    # Check if dataFim = None and adjusts correct type
    if strDataFim == "nan":
        strDataFim = None
    else:
        # Adjusts to ddmmYYYY
        "".join(strDataFim.split("-")).lstrip("0")

    # If current Environment is Homol, grab Homol ID from Excel (Homol ID not used in prod)
    if strEnvType == "HML":
        strHomolId = str(int(in_rowItem["HomolId"]))
    else:
        strHomolId = None

    return strAgencia, strConta, strDataInicio, strDataFim, strHomolId

# Requests Auth Token API and returns Access_Token to use in the next API and the date it will expire.
# Input: Config Dictionary
def request_access_token(in_dictConfig):
    # Get Environment Type from Config File (HML or PRD)
    strEnvType = in_dictConfig["EnvType"]

    # Check Environment and grab info according to it
    if strEnvType.upper() == "HML":
        strUrl = in_dictConfig["AuthTokenHomolURL"]
        strClientId = in_dictConfig["ClientIdHomol"]
        strClientSecret = in_dictConfig["ClientSecretHomol"]
    elif strEnvType.upper() == "PRD":
        strUrl = in_dictConfig["AuthTokenURL"]
        strClientId = in_dictConfig["ClientId"]
        strClientSecret = in_dictConfig["ClientSecret"]
    else:
        raise Exception("Environment Type not Mapped.")

    # Combine and encode credentials to base64
    strCredentials = f"{strClientId}:{strClientSecret}"
    strEncodedCredentials = base64.b64encode(strCredentials.encode("utf-8")).decode("utf-8")

    # Initialize and set Header Settings
    dictHeaders = {
        "Authorization": f"Basic {strEncodedCredentials}",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    # Initialize and set Body Settings
    dictPayloadData = {
        "grant_type": "client_credentials",
        "scope": in_dictConfig["ScopeToBeRequested"]
    }

    # Request Auth Token API
    jsonResponse = requests.post(strUrl, headers=dictHeaders, data=dictPayloadData)

    # Checks API Result
    if jsonResponse.status_code == 200:
        #If Success loads Json to get access_token
        dictJsonResponse = json.loads(jsonResponse.text)
        strAccessToken = dictJsonResponse["access_token"]
        # Get the date it will expires
        strAccessTokenTimeout = dictJsonResponse["expires_in"]
        dateTimeDelta = dt.timedelta(seconds=strAccessTokenTimeout)
        dateAccessTokenExpires = dt.datetime.now() + dateTimeDelta
        # Return access token and the date the current token will expire
        return strAccessToken, dateAccessTokenExpires
    else:
        strStatusCode = str(jsonResponse.status_code)
        strReason = str(jsonResponse.reason)
        dictJsonResponse = json.loads(jsonResponse.text)
        strErrorMessage = ""
        try:
            # Try to get the most commom error structure for this API
            strErrorMessage = dictJsonResponse["error_description"]
        except:
            pass
        raise Exception(f"Error requesting access token. Status Code: {strStatusCode}. Reason: {strReason}. Error Message: {strErrorMessage}.")

# Request Extract API based on values from Input Excel File
# Input: Config Dictionary, Access_Token, Agencia, Conta, HomolId, DataInicio, DataFim, NumeroPagina e MaxQtdRegistro
# Returns List of transactions for the specified account
def request_extract_info(in_dictConfig, in_strAcessToken,in_strAgencia, in_strConta, in_strHomolId=None,
                         in_strDataInicioSolicitacao=None, in_strDataFimSolicitacao=None,
                         in_strNumeroPaginaSolicitacao=None, in_strQtdRegistroPaginaSolicitacao=None):
    # Get Environment Type from Config File (HML or PRD)
    strEnvType = in_dictConfig["EnvType"]
    intQtdTotalPagina = 1
    listLancamento = []

    # Check Environment and grab info according to it
    if strEnvType.upper() == "HML":
        strUrl = in_dictConfig["ExtractHomolURL"]
        strDevAppKey = in_dictConfig["DevAppKeyHomol"]
    elif strEnvType.upper() == "PRD":
        strUrl = in_dictConfig["ExtractURL"]
        strDevAppKey = in_dictConfig["DevAppKey"]
    else:
        raise Exception("Environment Type not Mapped.")

    # Build full request URL
    strUrl = "/".join([strUrl, "conta-corrente", "agencia", in_strAgencia, "conta", in_strConta])

    # Initialize and set Header Settings
    dictHeaders = {
        "Authorization": f"Bearer {in_strAcessToken}",
        "Content-Type": "application/json",
        "x-br-com-bb-ipa-mciteste": in_strHomolId
    }

    # Remove Homol Key if in Production Environment
    if strEnvType.upper() == "PRD":
        del dictHeaders["x-br-com-bb-ipa-mciteste"]

    # Check to get all Transactions from the account
    while int(in_strNumeroPaginaSolicitacao) <= intQtdTotalPagina:
        # Initialize and set Params Settings
        dictParams = {
            "gw-dev-app-key": strDevAppKey,
            "dataInicioSolicitacao": in_strDataInicioSolicitacao,
            "dataFimSolicitacao": in_strDataFimSolicitacao,
            "numeroPaginaSolicitacao": in_strNumeroPaginaSolicitacao,
            "quantidadeRegistroPaginaSolicitacao": in_strQtdRegistroPaginaSolicitacao
        }
        # Request Extract API
        jsonResponse = requests.get(strUrl, headers=dictHeaders, params=dictParams)
        # Adds +1 to request page in case there are more than one page
        in_strNumeroPaginaSolicitacao = str(int(in_strNumeroPaginaSolicitacao) + 1)
        # Check API Result
        if jsonResponse.status_code == 200:
            # If Success load Json
            dictJsonResponse = json.loads(jsonResponse.text)
            # Update Amount of Pages based on request
            intQtdTotalPagina = dictJsonResponse["quantidadeTotalPagina"]
            # Adds Transactions to listLancamento
            listLancamento.extend(dictJsonResponse["listaLancamento"])
        else:
            strStatusCode = jsonResponse.status_code
            strReason = jsonResponse.reason
            dictJsonResponse = json.loads(jsonResponse.text)
            strErrorMessage = ""
            try:
                #Try to get the most commom error structure for this API
                strErrorMessage = dictJsonResponse["message"]
            except:
                pass
            raise Exception(f"Error requesting Extract Info. Status Code: {strStatusCode}. Reason: {strReason}. Error Message: {strErrorMessage}.")
    # Return List of Transactions for specific account
    return listLancamento

# Convert list of Transactions to a DataFrame format to input in the Excel Report
# Input: List of Transactions
# Returns dataFrame ready to be inputted into Excel
def convert_extract_info_to_df(in_listLancamento):
    # Initialize All lists to Write into Output Excel
    listDataLancamento = []
    listNumeroDocumento = []
    listValorLancamento = []
    listTextoDescricaoHistorico = []

    # Gather all necessary info to write into excel
    for item in in_listLancamento:
        # Get wanted fields from List of Extract
        strDataLancamento = str(item["dataLancamento"])
        strNumeroDocumento = str(item["numeroDocumento"])
        fltValorLancamento = item["valorLancamento"]
        strTextoDescricaoHistorico = item["textoDescricaoHistorico"]

        # Append Items To List to add to Excel Report
        listDataLancamento.append(strDataLancamento)
        listNumeroDocumento.append(strNumeroDocumento)
        listValorLancamento.append(fltValorLancamento)
        listTextoDescricaoHistorico.append(strTextoDescricaoHistorico)

    # Build data dictionary to convert to dataframe
    data = {
        'dataLancamento': listDataLancamento,
        'numeroDocumento': listNumeroDocumento,
        'valorLancamento': listValorLancamento,
        'textoDescricaoHistorico': listTextoDescricaoHistorico
    }
    # Returns dataframe with all transactions
    return pd.DataFrame(data)

#Converts Lists of logs to dataframe format to input in the Excel Report
# Input: List of Agencia, List of Conta, List of Status
# Returns dataFrame ready to be inputted into Excel
def convert_log_info_to_df(in_listLogsAgencia, in_listLogsConta, in_listLogsStatus):
    # Build data dictionary to convert to dataframe
    data = {
        'Agencia': in_listLogsAgencia,
        'Conta': in_listLogsConta,
        'Status': in_listLogsStatus
    }
    # Returns dataframe with all logs
    return pd.DataFrame(data)

# Create and Write Excel file if the file doesn't exist, append sheet to file if Excel already exists
# Input: Config Dictionary, dataFrame which will be inserted, Agencia, Conta and boolLogs True if adding logs to Excel
def write_excel(in_dictConfig, in_dfToInsert, in_strAgencia, in_strConta, boolLogs):
    # Sets Output Excel File Path
    strExcelFileName = in_dictConfig["OutputExcelFileName"]
    strOutputFolder = in_dictConfig["OutputFolder"]
    strExcelFilePath = os.path.join(os.curdir, strOutputFolder, strExcelFileName)

    # Sets Sheet Name
    if not boolLogs:
        #SheetName = 'Agencia - Conta'
        strSheetName = "-".join([in_strAgencia, in_strConta])
    else:
        #SheetName = Name specified in the Config File
        strSheetName = in_dictConfig["OutputLogsSheetName"]

    # Checks if Excel already exists
    if not os.path.exists(strExcelFilePath):
        # Create Excel and write all rows from extract into new sheet
        with pd.ExcelWriter(strExcelFilePath) as writer:
            in_dfToInsert.to_excel(writer, sheet_name=strSheetName, index=False)
    else:
        # Append all rows from extract to existing Excel
        with pd.ExcelWriter(strExcelFilePath, mode="a") as writer:
            in_dfToInsert.to_excel(writer, sheet_name=strSheetName, index=False)