import duckdb

query = """
DESCRIBE
SELECT *
FROM read_xlsx('data.xlsx')
"""

print(duckdb.sql(query))