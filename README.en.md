# 📊 Bank Statement Extraction Automation (Python)

This project automates extract data extraction
banking via API, processing multiple accounts from one file
Excel and generating structured output reports.

The solution includes authentication, data paging, processing of
errors and log generation, being ideal for automation scenarios
finance and RPA.

------------------------------------------------------------------------

## 🚀 Main Features

- 🔐 Automatic authentication via API (OAuth / Client Credentials)
- 📥 Reading accounts from Excel file
- 🔄 Automatic pagination when collecting data
- 📊 Transformation of data into DataFrame (pandas)
- 📁 Excel report generation per account
- 🧾 Record execution logs (success and error)
- ⚙️ Support for multiple environments (HML / PRD)

------------------------------------------------------------------------

## 🧱 Project Structure 

    . 
    ├── bot.py # Main script (orchestration) 
    ├── CustomLib.py # Auxiliary functions and business rules 
    ├── Config/ 
    │ └── Config.env # Process Settings 
    ├── Input/ 
    │ └── input.xlsx # Input file 
    ├── Output/ 
    │ └── output.xlsx # Generated file

------------------------------------------------------------------------

## ⚙️ Configuration

The project uses a `.env` file to centralize configurations.

------------------------------------------------------------------------

## 📥 Input File Format

The input Excel must contain the following columns:

| Field | Description
| ------------ | -----------------------------
| Agency | Agency number
| Account | Account number
| Start Date | Start date (optional)
| End Date | End date (optional)
| HomolId | Required for HML only

------------------------------------------------------------------------

## ▶️ Execution

### 1. Install dependencies 

    pip install -r requirements.txt

### 2. Configure the `.env` file 

    Fill in the API credentials.

### 3. Execute the process 

    python bot.py

------------------------------------------------------------------------

## 🧠 Process Flow

1. Reading the configuration file (`.env`)
2. Reading Input Excel
3. For each account: 
    -   Adjustment of data (agency, account, dates) 
    -   Generation/validation of the access token 
    -   Extract API consumption (with pagination) 
    -   Conversion of data to DataFrame 
    -   Writing in Excel output
4. Generation of log tab with status per account

------------------------------------------------------------------------

## 📊 Generated Output

The final Excel file contains:

- 📄 One tab per account (`Agencia-Conta`)
- 🧾 **Logs** tab with execution status

------------------------------------------------------------------------

## ⚠️ Error Handling

- API failures are captured and logged
- Processing continues even in the event of an error on an account
- Error messages are stored in the logs tab

------------------------------------------------------------------------

## 🛠️ Technologies Used

- Python 3.x\
- pandas\
- requests\
- python-dotenv\
- openpyxl

------------------------------------------------------------------------

## 👤 Author

Pedro Coletti
