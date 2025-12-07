#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''This script connects to a PostgreSQL database, make a query by using a python variable,
retrieves data from a table, and appends new data to it.

To create the table, use the following SQL command:
CREATE TABLE table_1 (
    id INTEGER,
    name varchar(50)
);

To insert data into the table:
INSERT INTO table_1 VALUES (1,'Juan');

To delete data from the table:
DELETE FROM table_1 WHERE id = 1;

To delete the table:
DROP TABLE table_1;
'''


# Import libraries
import pandas as pd
from postgreSQL_sqlalchemy_example import create_engine

# Define the parameters to connect to the database
db_config = {'user': 'tony',
             'pwd': '1234',
             'host': 'localhost',
             'port': 5432,
             'db': 'db_1'
             }
# username, password, server address, port, database name

# Create database connection string .
connection_string = 'postgresql://{}:{}@{}:{}/{}'.format(db_config['user'],
                                                         db_config['pwd'],
                                                         db_config['host'],
                                                         db_config['port'],
                                                         db_config['db'])
# Connect to the database.
engine = create_engine(connection_string)

# Create a SQL query
value = 1
query = ''' SELECT *
            FROM table_1
            WHERE id = '{}'
        '''.format(value)

# Obtain data from query and store it in a DataFrame
df = pd.io.sql.read_sql(query, engine)

# Show the first rows of the DataFrame
print(df.head())

# Create DataFrame form dictionary
df2 = pd.DataFrame(
    {
        'id': [2, 3, 4, 5, 6, 7, 8, 9],
        'name': ['Antonio', 'Pedro', 'Maria', 'Luis', 'Ana', 'Jorge', 'Sofia', 'Diego']
    })

print('New data')
print(df2)

# Append new data to the table. The parameter 'if_exists' specifies what to do if the table already exists.
# if_exists={'append', 'replace', 'fail'}.
# Use 'fail' to raise an error if the table exists. If table does not exist, it will be created
df2.to_sql(name='table_1', con=engine, if_exists='append', index=False)
