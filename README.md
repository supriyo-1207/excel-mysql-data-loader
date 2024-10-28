# Excel to MySQL Loader

This project is a Python script that reads data from an Excel file and inserts it into a MySQL database. The script handles multiple sheets, creates tables dynamically based on the sheet names, and inserts data into these tables. It also includes validation to check for the existence of tables and columns before creating or inserting data.

## Features

- Load data from an Excel file with multiple sheets.
- Dynamically create MySQL tables based on sheet names.
- Insert data into MySQL tables, handling data types appropriately.
- Validate if data already exists before inserting new data.

## Technologies Used

- **Python**: Programming language used for scripting.
- **Pandas**: Library for data manipulation and analysis.
- **openpyxl**: Library to read Excel files.
- **mysql-connector-python**: MySQL driver for Python to interact with MySQL databases.
- **python-dotenv**: Library to manage environment variables.

## Requirements

- Python 3.x
- Pandas
- openpyxl
- mysql-connector-python
- python-dotenv

## Setup

1. **Clone the repository**:
    ```sh
    git clone https://github.com/SupriyoMaity1207/excel-mysql-data-loader.git
    ```

2. **Create a virtual environment and activate it**:
    ```bash
    python -m venv venv
    ```

3. **Activate the virtual environment**:
    - On Windows:
      ```bash
      .venv\Scripts\activate
      ```
    - On macOS/Linux:
      ```bash
      source venv/bin/activate
      ```

4. **Install the required packages**:
    ```bash
    pip install -r requirements.txt
    ```

5. **Create a `.env` file in the root directory and add your configuration**:
    ```bash
    FILE_PATH=path_to_your_excel_file.xlsx
    DB_NAME=your_database_name
    DB_PASSWORD=your_database_password
    ```

## Usage

Run the script:
```bash
python your_script_name.py