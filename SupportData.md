# SupportData

Loads all tables in a given file and applies necessary rename and type information before returning the tables in a named record format for easy access in code.

- This code supersedes the LoadSupportData which was inefficient and constrained to Excel files.
- The code uses buffering to reduce file access.
- It supports zip file with multiple csv tables.

## Parameters

| Parameter | Description |
|-----------|-------------|
| Path      | Support Path. Defaults to P[Support Path].|
| File Name | Partial or full name of Support File. Defaults to Support.|

## Performance

- Averages of 50 iterations
  - 0.00:00:00.4384429 ? Support Excel file with 51 tables of size 3.46MB from shared folder using old code.
  - 0.00:00:00.0196887 ? Support Zip file with 51 csv files of size 801KB from shared folder using new code.
- Speedup of 19 to 22 times.
- Huge reduction in data transfer. 173MB compared to 801KB due to efficient buffering.
