import pandas as pd
import openpyxl
import mysql.connector
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Path to Excel file
excel_file = os.environ.get('FILE_PATH')
print(excel_file)

# Read all sheets into a dictionary of dataframes
sheets_dict = pd.read_excel(excel_file, sheet_name=None)

# Establish the connection to MySQL
db_connection = mysql.connector.connect(
    host='localhost',
    user='root',
)

curser = db_connection.cursor()

if db_connection.is_connected():
    print("Connected to MySQL database")
else:
    print("Failed to connect to MySQL database")

# Get the database name from the environment variable
# db_name = os.environ.get('DB_NAME')
db_name="insert_Excel_data_in_Mysql"
print(db_name)

# Create the database if it doesn't exist
curser.execute(f"CREATE DATABASE IF NOT EXISTS {db_name}")
db_connection.close()

# Reconnect to the database after creation
db_connection = mysql.connector.connect(
    host='localhost',
    user='root',
    database=db_name
)

curser = db_connection.cursor()

# Map pandas dtypes to MySQL column types
dtype_mapping = {
    'int64': 'INT',
    'float64': 'FLOAT',
    'object': 'VARCHAR(255)',  # For strings
    'bool': 'BOOLEAN',
    'datetime64[ns]': 'DATETIME'
}

# Flag to track if any duplicates are found
duplicates_found = False

# Iterate through the sheets and create tables
for sheet_name, df in sheets_dict.items():
    print(sheet_name)
    print(df.info())
    print(df.columns)

    # Replace NaN values with None
    df = df.where(pd.notnull(df), None)

    # Check if the table exists
    curser.execute(f"SHOW TABLES LIKE '{sheet_name}'")
    result = curser.fetchone()

    if result:
        print(f"Table `{sheet_name}` already exists.")
    else:
        # Create table if it doesn't exist
        columns = []
        for column_name, dtype in zip(df.columns, df.dtypes):
            mysql_dtype = dtype_mapping.get(str(dtype), 'VARCHAR(255)')
            columns.append(f"`{column_name}` {mysql_dtype}")

        column_definition = ", ".join(columns)
        create_table = f"CREATE TABLE `{sheet_name}` ({column_definition})"
        
        # Print the generated SQL for debugging
        print(create_table)
        
        # Execute the table creation statement
        curser.execute(create_table)

    # Insert rows into the table
    insert_columns = ", ".join([f"`{col}`" for col in df.columns])
    placeholders = ", ".join(["%s"] * len(df.columns))
    insert_query = f"INSERT INTO `{sheet_name}` ({insert_columns}) VALUES ({placeholders})"
    
    # Determine unique columns (for simplicity, assuming the first column is unique)
    unique_column = df.columns[0]  # Adjust this as needed
    
    # Check if data already exists before inserting
    for index, row in df.iterrows():
        row_data = tuple(row)  # Convert the row to a tuple
        
        # Check if the row data already exists in the database
        select_query = f"SELECT COUNT(*) FROM `{sheet_name}` WHERE `{unique_column}` = %s"
        curser.execute(select_query, (row[unique_column],))
        count = curser.fetchone()[0]
        
        if count > 0:
            duplicates_found = True
            break  # Exit loop after finding the first duplicate

    if duplicates_found:
        print(f"Duplicate data found in table `{sheet_name}`. No new data will be inserted.")
    else:
        for index, row in df.iterrows():
            row_data = tuple(row)  # Convert the row to a tuple
            curser.execute(insert_query, row_data)

# Commit changes and close the connection
db_connection.commit()
curser.close()
db_connection.close()
