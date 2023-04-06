# README
## Description
This Python script, main.py, is designed to read data from an Excel workbook (Teste_Entradas.xlsx) using the openpyxl library, perform calculations, and update a summary worksheet in the same workbook with the calculated values. The script calculates the total for each month by summing the values from different categories in different sheets of the workbook, and writes the calculated totals to the corresponding cells in the summary worksheet.

## Requirements
- Python 3.x
    - openpyxl library (can be installed using pip install openpyxl)
## Usage
1. Ensure that the Teste_Entradas.xlsx file is in the same directory as the main.py script.
2. Run the main.py script using a Python interpreter or an Integrated Development Environment (IDE).
3. The script will read the data from the Teste_Entradas.xlsx file, perform calculations, and update the summary worksheet with the calculated totals for each month.
4. The updated data will be saved to a new file called entradas.xlsx in the same directory as the Teste_Entradas.xlsx file.

## Customization
- You can customize the sheet names and column indices for each category by modifying the categories list in the script. Each category should be represented as a tuple containing the sheet name and the column index.
- You can customize the month names and corresponding column indices by modifying the months list in the script. Each month should be represented as a tuple containing the month name and the column index in the summary worksheet.

## Note
1. The script assumes that the input workbook (Teste_Entradas.xlsx) has the following sheets with the specified names: CCLops, CCProd, CCProj, and Resumo.
2. The script uses the sum() function and generator expressions to calculate the totals for each month, and the format() function to format the calculated totals as currency values with two decimal places.
3. The calculated totals are written to the summary worksheet (Resumo) in the same workbook (Teste_Entradas.xlsx) in the cells specified in the script.
4. It's recommended to make a backup of the input workbook before running the script to avoid data loss.