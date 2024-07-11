# ExcelToJsonWizard
A tool for converting Excel data to JSON files and generating C# loader classes for use in Unity.

# JSON Converter Tool

This tool converts Excel files to JSON files and generates corresponding C# loader classes. The tool is configured using a `config.txt` file and supports various options, including handling multiple sheets and using Unity's Resources folder for JSON file loading.

## Features
- Converts Excel files to JSON files.
- Generates C# loader classes for the JSON files.
- Supports Enum definitions from a separate `Enum.xlsx` file.
- Configurable via a `config.txt` file.
- Handles multiple sheets in Excel files.
- Supports Unity's `Resources` folder for loading JSON files.

## Prerequisites
- .NET Core or .NET Framework installed.
- ClosedXML library for reading Excel files.

## Installation
1. Clone the repository.
2. Open the solution in your preferred IDE.
3. Restore NuGet packages.

## Configuration
The tool uses a `config.txt` file for configuration. The `config.txt` file and necessary directories will be automatically generated on the first run.

## Usage
1. **Prepare your Excel files**:
   - Place your Excel files in the directory specified in `config.txt`.
   - Ensure your Excel files are formatted correctly (see below).

2. **Run the tool**:
   - Execute the compiled program.
   - The tool will read the Excel files, generate JSON files and C# loader classes, and save them in the specified directories.

## Excel File Format
- The first row should contain variable names.
- The second row should contain data types.
- The third row should contain descriptions (optional, but "No description provided." will be used if empty).
- The fourth row and beyond should contain the data.

## Enum Definitions
- Enum definitions should be placed in the first sheet of `Enum.xlsx`.
- The first row should contain the Enum name.
- The second column onwards should contain the Enum values.
- Enum definitions are used from the first row onwards without a header row.
- If Enum definitions are not needed, there is no need to create the `Enum.xlsx` file.

## Possible Errors and Troubleshooting
1. **Missing `key` column**:
   - Ensure the first column in your Excel file is named `key`.

2. **Incorrect data type**:
   - Ensure the data type specified in the second row is correct and supported (e.g., `int`, `string`, `List<int>`, etc.).

3. **Duplicate `key` values**:
   - Ensure that there are no duplicate values in the `key` column.

4. **Enum type not found**:
   - Ensure the `Enum.xlsx` file exists and is formatted correctly.
   - Ensure the Enum name in the Excel file matches the name in `Enum.xlsx`.

5. **Resource loading issues**:
   - If using Unity's `Resources.Load`, ensure the path is correct and the file exists in the Resources folder.
