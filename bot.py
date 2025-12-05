import os
import logging as log
import datetime as dt
import CaseCSNlib as customLib

def main():
    # Set Logging Level and Format
    log.basicConfig(level=log.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Initiallize and Set Global Variables
    strConfigFileName = "Config.env"
    strConfigFilePath = "\\".join([os.curdir,"Config",strConfigFileName])
    boolValidToken = False
    dateAccessTokenExpires = None
    strAccessToken = None
    listLogsAgencia = []
    listLogsConta = []
    listLogsStatus = []

    # Read Config Files and returns the dictionary used throughout the process
    dictConfig = customLib.read_config(strConfigFilePath)
    log.info("Config file read sucessfully.")

    # Read the Input Excel File based on Config settings and returns the dataframe of that Excel
    dfExcel = customLib.read_excel(dictConfig)
    log.info("Input Excel file read sucessfully.")

    #For each row in input excel get access token when necessary and request extract info and add to OutputExcel
    for index, row in dfExcel.iterrows():
        # Set and adjusts input excel variables to be used on the request_extract_info function
        strAgencia, strConta, strDataInicio, strDataFim, strHomolId = customLib.set_excel_variables(dictConfig, row)
        log.info(f"Started process for account {strConta}.")
        # Get Config data related to the request_extract_info function
        strPageNumber = dictConfig["ExtractPageNumber"]
        strMaxEntries = dictConfig["ExtractMaxEntries"]
        try:
            # Checks if Access_Token no longer valid
            if not boolValidToken:
                # Requests API Auth Token and returns Access_Token to use in the next API and the date it will expire
                strAccessToken,dateAccessTokenExpires = customLib.request_access_token(dictConfig)
                boolValidToken = True
                log.info("Access Token API request successful.")
            else:
                log.info("Access Token still valid.")

            # Request Extract API and return list of transactions
            listLancamento = customLib.request_extract_info(dictConfig, strAccessToken, strAgencia,
                                                            strConta, strHomolId, strDataInicio,
                                                            strDataFim, strPageNumber, strMaxEntries)
            log.info("Extract API request successful.")

            # Convert all lines from list of transactions to DataFrame
            dfToInsert = customLib.convert_extract_info_to_df(listLancamento)

            # Write Extract information in output Excel File
            customLib.write_excel(dictConfig, dfToInsert, strAgencia, strConta, boolLogs=False)
            log.info(f"Transaction items for account {strConta} inserted successfully in output file.")

            # Set Log info for Success
            listLogsAgencia.append(strAgencia)
            listLogsConta.append(strConta)
            listLogsStatus.append("OK")

            # Check if Access_Token is still Valid
            if dt.datetime.now() >= dateAccessTokenExpires:
                boolValidToken = False
        except Exception as e:
            # Set log info for Error
            listLogsAgencia.append(strAgencia)
            listLogsConta.append(strConta)
            listLogsStatus.append(str(e))
            log.warning(f"Exception occurred while processing account {strConta}. Error: {e}")

    #Convert List of Logs to Data Frame
    dfLogsToInsert = customLib.convert_log_info_to_df(listLogsAgencia, listLogsConta, listLogsStatus)
    # Add log to output Excel File
    customLib.write_excel(dictConfig, dfLogsToInsert, None, None, boolLogs=True)
    log.info("Logs add sucessfully to output file.")

if __name__ == "__main__":
    main()