#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File:           sqlalchemy_sqlserver_example.py
Author:         Antonio Arteaga
Last Updated:   2025-11-26
Version:        1.0
Description:
Simple example of connection to a DB by using SQLServer, sqlalchemy and Python.
The query result is saved in an excel file.
Dependencies:   pandas==2.3.3, openpyxl==3.1.5, SQLAlchemy==2.0.44.
Driver ODBC:  ODBC Driver 18 for SQL Server.
Usage:          The sql query has the form:
SELECT TOP 4 *
FROM "tabla1"
WHERE Columna1 IN ('valor1')
"""

import pandas as pd
from sqlalchemy import create_engine
import time

# Parameters to connect to the database
server = 'localhost\\SQLEXPRESS'
database = 'tabla1'
username = 'arteaga7'
password = 'pa$$Word'

engine = create_engine(
    f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=SQL+Server"
)


table = [
    "tabla1"
    "tabla2"
]
Columna1 = 'valor1'

query = f"""
SELECT TOP 4 *
FROM [{table[0]}]
WHERE Columna1 = '{Columna1}'
"""

t1 = time.time()
df = pd.read_sql(query, engine)
print("Tiempo de consulta (sqlalchemy):", time.time() - t1)
df.to_excel(f"{table[0]}.xlsx", index=False)
print(f"Tabla '{table[0]}' guardada con éxito!")
