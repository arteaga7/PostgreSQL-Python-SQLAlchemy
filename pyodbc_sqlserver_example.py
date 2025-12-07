#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File:           pyodbc_sqlserver_example.py
Author:         Antonio Arteaga
Last Updated:   2025-11-26
Version:        1.0
Description:
Simple example of connection to a DB by using SQLServer and Python.
The query result is saved in an excel file.
Dependencies:   pyodbc==5.3.0, pandas==2.3.3, openpyxl==3.1.5.
Driver ODBC:  ODBC Driver 18 for SQL Server.
Usage:          The sql query has the form:
SELECT TOP 4 *
FROM "tabla1"
WHERE Columna1 IN ('valor1')
"""

import pyodbc
import pandas as pd
import time

# Connection parameters
server = 'localhost\\SQLEXPRESS'
database = 'tabla1'
username = 'arteaga7'
password = 'pa$$Word'

# Connection configuration
connection_string = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    "Encrypt=no;"              # o "Encrypt=yes;TrustServerCertificate=yes;"
)

table = ["tabla1"]
Columna1 = ["valor1"]
# Create placeholders: "?, ?"
placeholders = ", ".join("?" * len(Columna1))

query = f"""
SELECT TOP 4 *
FROM [{table[0]}]
WHERE Columna1 IN ({placeholders})
"""
params = Columna1

try:
    conn = pyodbc.connect(connection_string)
    t1 = time.time()
    df = pd.read_sql(query, conn, params=Columna1)
    print("Tiempo de consulta (pyodbc):", time.time() - t1)
    df.to_excel(f"{table}.xlsx", index=False)
    print(f"Tabla '{table}' guardada con éxito!")
    conn.close()

except Exception as e:
    print("Error al conectar o ejecutar consulta:", e)
