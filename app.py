import streamlit as st
import pyodbc
import pandas as pd

# Database Connection Function
def connect_db():
    server = r'DESKTOP-2OEF9G3\SQLEXPRESS'  # Change if needed
    database = 'StreamlitDB'

    try:
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={server};",
            autocommit=True
        )
        cursor = conn.cursor()
        cursor.execute(f"IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = '{database}') CREATE DATABASE {database}")
        conn.close()

        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};",
            autocommit=True
        )
        return conn
    except Exception as e:
        st.error(f"Database connection failed: {e}")
        return None

# Table Creation Function
def create_table():
    conn = connect_db()
    if conn:
        cursor = conn.cursor()
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Users')
            CREATE TABLE Users (
                ID INT IDENTITY(1,1) PRIMARY KEY,
                Name NVARCHAR(100),
                Age INT
            )
        """)
        conn.close()

# Insert Data Function
def insert_data(name, age):
    conn = connect_db()
    if conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Users (Name, Age) VALUES (?, ?)", (name, age))
        conn.close()
        st.success("Data inserted successfully!")

# Fetch Data Function
def fetch_data():
    conn = connect_db()
    if conn:
        df = pd.read_sql("SELECT * FROM Users", conn)
        conn.close()
        return df
    return pd.DataFrame()

# Download Data as Excel
def download_excel():
    df = fetch_data()
    if not df.empty:
        excel_file = "User_Data.xlsx"
        df.to_excel(excel_file, index=False)
        return excel_file
    return None

# Streamlit UI
st.title("SQL Server CRUD with Streamlit")

# Create table
create_table()

# Insert Data Section
st.subheader("Insert New User")
name = st.text_input("Enter Name")
age = st.number_input("Enter Age", min_value=1, step=1)
if st.button("Insert Data"):
    insert_data(name, age)

# Show Data Section
st.subheader("User Records")
df = fetch_data()
st.dataframe(df)

# Download Data Section
if not df.empty:
    excel_file = download_excel()
    with open(excel_file, "rb") as file:
        st.download_button("Download Data as Excel", file, file_name="User_Data.xlsx")

