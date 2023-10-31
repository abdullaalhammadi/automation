# Documentation: SQL to Excel Data Extraction

This documentation explains a Python script designed to read SQL queries from a JSON file, execute them on a CSV file using the DuckDB engine, convert the results to pandas DataFrames, and then write those results into specified ranges in an Excel file.

## 1. Required Libraries:

- `pandas`: Used for data manipulation and exporting data to Excel.
- `duckdb`: An embedded analytical data management system. It is used here to run SQL queries.
- `string`: Standard Python library used for string operations.
- `time`: Standard python library for time-based operations.
- `tqdm`: Loop progression tracker library used in python.
- `re`: Standard Python library for regular expression operations.

## 2. Input Data:

### 2.1. CSV File (`merged_gw.csv`):

This file contains the primary data which the SQL queries will run against. In the script, any column named "kickoff_time" is dropped before further processing.

### 2.2. JSON File (`queries.json`):

This file contains a list of SQL queries and related configurations. Each entry has:

- `indicator_number`: A unique identifier for the query.
- `query`: The SQL query to run.
- `range`: The target range in Excel where the result of the query will be written.
- `startrow`: The starting row in Excel for the result.
- `startcol`: The starting column in Excel for the result.

## 3. Output:

The output is an Excel file named `FPL.xlsx`. If the file already exists, the script appends data to it.

## 4. Script Workflow:

1. Read the `merged_gw.csv` file into a pandas DataFrame and drop the "kickoff_time" column.
2. Read the `queries.json` file into a pandas DataFrame named `indicators`.
3. Open (or create) the `FPL.xlsx` Excel file.
4. Loop over each entry in the `indicators` DataFrame.
5. Convert the `startcol` (Excel-style column notation) to a zero-based numeric column index.
6. Run the SQL query using the DuckDB engine and convert the result to a pandas DataFrame.
7. Write the result into the Excel file at the specified starting row and column.

## 5. Understanding the Column Conversion:

For converting Excel-style column notation to a zero-based numeric column index, the script uses a loop. For example, the column "B" corresponds to index 1. This conversion is essential for writing data to Excel in the correct location.

## 6. Important Notes:

- The Excel engine used is `openpyxl`.
- If the sheet already exists in the Excel file, the script will overlay (overwrite) the data.

## Example:

Consider an entry in `queries.json`:

```json
{
    "indicator_number": "A1",
    "query": "...",
    "range": "B4:J23"
}
```

For this entry, the SQL query will run, and its result will start populating in the Excel file from row 4 and column B.

