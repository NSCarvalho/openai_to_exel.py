## Project Overview
This project is a Python script designed to parse OpenAPI specification files (in JSON format) and generate an Excel file that lists all request body fields, their types, required status, and descriptions. Each API endpoint is represented by a separate sheet, and the sheet name is the first 30 characters of the `summary`. Additionally, before the fields, a summary table with the `endpoint`, `method`, and `summary` is included for easy reference.

This project was developed with the assistance of an AI model to enhance the development process and improve efficiency.

## Features
- Supports complex request bodies with nested objects and arrays.
- Handles `$ref` references in OpenAPI schemas.
- Creates a separate sheet for each service in the OpenAPI document.
- Provides a summary of each service in the first row of each sheet, including `endpoint`, `method`, and `summary`.
- Adds tables listing the fields of the request body, with detailed information on field names, types, required status, and descriptions.

## Installation

1. **Clone the repository** (or download the script file directly).

2. **Ensure Python is installed**:
   This script is compatible with Python 3.x. You can download Python from [python.org](https://www.python.org/downloads/) if you haven't installed it yet.

3. **Run the script**:
   The script automatically handles the installation of the necessary dependencies. Simply run the following command:

   ```bash
   python openapi_to_excel.py

## Usage

1. **Prepare your OpenAPI file**:
   - Make sure you have an OpenAPI specification file in JSON format. This file will be parsed by the script to generate the Excel output.
   - The OpenAPI file should follow the standard OpenAPI 3.0+ format.

2. **Running the script**:
   To run the script, use the following command:

   ```bash
   python openapi_to_excel.py
   ```

   By default, the script expects the OpenAPI file to be named `openapi.json` and will generate an Excel file named `openapi_as_exel.xlsx`.

3. **Customize file paths**:
   If your OpenAPI file has a different name or location, or if you'd like to save the Excel file with a different name, modify the `openapi_path` and `excel_path` variables in the script:

   ```python
   openapi_path = "path/to/your/openapi.json"
   excel_path = "path/to/output/openapi_as_exel.xlsx"
   ```

4. **Resulting Excel file**:
   After running the script, you will find the generated Excel file (`openapi_as_exel.xlsx`) in the output directory. The Excel file will contain:
   - A separate sheet for each endpoint, named as summary limited to 30 characters.
   - In each sheet, the first rows will display the `endpoint`, `method`, and `summary` of the service.
   - Below that, a table will list all the fields in the request body, showing their names, types, whether they are required, and their descriptions.

## Example Output

Each sheet in the generated Excel file will contain:
1. **Service Information**:
   - Endpoint: `/users/{id}`
   - Method: `GET`
   - Summary: `Retrieve user by ID`

2. **Request Body Fields**:
   - Field Name: `userId`
   - Type: `integer`
   - Required: `Yes`
   - Description: `The unique ID of the user.`

## Contribution

Feel free to fork the repository, submit issues, and create pull requests. Contributions to improve the script or add additional features are welcome.

## License

This project is licensed under the MIT License.
