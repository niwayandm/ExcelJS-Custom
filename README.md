# Custom ExcelJS

## Overview

This repository contains custom ExcelJS functions to facilitate exporting HTML tables to Excel with automatic merging and styling. The functions handle various aspects like column address computation, cell styling, merging cells, and handling special cell formats (e.g., percentages and numbers).

## Features

- **Multiple Table Export in 1 sheet**: Export multiple HTML tables into 1 sheet by providing custom table classes. 
- **Automatic Cell Merging**: Automatically merge cells based on HTML table structure.
- **Custom Cell Styling**: Apply background color, font, alignment, and borders to cells.
- **Special Formatting**: Handle percentage and numeric formats accurately.
- **Grouping and Hiding Rows**: Support for grouping rows and setting outline levels.

## Installation

### Clone the Repository
```bash
git clone https://github.com/niwayandm/ExcelJS-Custom.git
```

### Include in Your Project
```html
<script src="path/to/CustomExport.js"></script>
```

Ensure you have ExcelJS included in your project as well:
```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.2.0/exceljs.min.js"></script>
```

## Usage
### Example
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Custom ExcelJS Export Example</title>
</head>
<body>
    <table class="exportTable">
        <tr>
            <td>Item</td>
            <td>Price</td>
        </tr>
        <tr>
            <td>Apple</td>
            <td>1.00</td>
        </tr>
        <tr>
            <td>Banana</td>
            <td>0.50</td>
        </tr>
    </table>
    <button onclick="exportToExcel()">Export to Excel</button>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.2.0/exceljs.min.js"></script>
    <script src="path/to/CustomExport.js"></script>
    <script>
        function exportToExcel() {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Sheet 1');

            exportTableToWorksheet('exportTable', worksheet);

           try {
                var title = `Export Excel Table`
                await new Promise(resolve => setTimeout(resolve, 500)); // Delay 
                workbook.xlsx.writeBuffer().then(function(buffer) {
                    let blob = new Blob([buffer], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    });
                    let link = document.createElement('a');
                    link.href = window.URL.createObjectURL(blob);
                    link.download = title.toUpperCase() + '.xlsx';
                    link.click();
                });

                } catch (error) {
                console.error('Error:', error);
                } finally {
                $('#loaderExcel').hide();

            }
        }
    </script>
</body>
</html>
```

## Functions
- **delay(ms)**: Returns a promise that resolves after the specified milliseconds.
- **getColumnAddress(colNum)**: Converts a column number to its corresponding Excel column address (e.g., 1 -> A, 27 -> AA).
- **setCellStyle(htmlCell, excelCell)**: Applies styles from an HTML cell to an ExcelJS cell.
- **getColSpan(worksheet, cellAddress, rowNumber)**: Gets the colspan of a cell in the worksheet.
- **getMaxRow(worksheet)**: Returns the maximum row number in the worksheet.
- **isTopLeftCellOfMerge(worksheet, cellAddress)**: Checks if the cell is the top-left cell of a merged cell range.
- **columnToNumber(column)**: Converts an Excel column address to its corresponding column number.
- **extractTextFromHtml(html)**: Extracts text content from an HTML string.
- **exportTableToWorksheet(tableClassName, worksheet, startRow)**: Exports an HTML table to an ExcelJS worksheet.

## Acknowledgments
- Thanks to the ExcelJS team for their powerful library.
