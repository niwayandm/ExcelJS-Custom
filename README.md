# Custom ExcelJS

## Overview

This repository contains my custom ExcelJS functions to facilitate exporting HTML tables to Excel with styling and merged cell support directly from HTML structure and CSS.

## Features

### Table & Cell Handling
- **Multiple Table Export in One Sheet**  
  Export multiple HTML tables to a single worksheet by providing a class name.
  
- **Automatic Cell Merging**  
  Handles `rowspan` and `colspan` using ExcelJS cell merging.

- **Dynamic Column Widths**  
  Automatically calculates and sets appropriate column widths based on content.

- **Text Wrapping & Line Breaks**  
  Supports `<br>` tags by converting them to `\n` and enabling `wrapText`.

### Styling & Formatting
- **Custom Cell Styling**  
  Applies styles such as:
  - Background color (cascades from cell → row → section)
  - Font color and size
  - Bold (from computed style or presence of `<b>` / `<strong>`)
  - Vertical/horizontal alignment (`.text-center`, `.text-right`)
  - Text rotation (from CSS `transform` or `writing-mode`)
  - Borders for all cells

- **Partial Bold Support**  
  Inline `<b>` and `<strong>` tags inside cells are rendered as bold using ExcelJS `richText`, without overriding existing styles.

- **Number & Percentage Formatting**  
  - Detects numeric values and applies comma-separated formatting (`#,##0.00`)
  - Automatically formats percentages (e.g., `25%` becomes `0.25` with Excel `%` format)

### Row & Column Controls
- **Row Grouping & Collapse Support**  
  Use HTML classes to control outline/grouping levels:
  - `<tr class="collapse">` → grouped and hidden
  - `<tr class="collapse in">` → grouped and visible
  - `<tr class="collapse-level-2">` → grouped at level 2
  - `<tr class="hidden">` → row will be hidden

- **Hidden Columns Based on Headers**  
  - Add `.hide-column` to any `<th>` element in `<thead>`  
  - Supports complex multi-row headers with `colspan` and `rowspan`


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
### Basic Example
```html
<table class="exportTable">
  <thead>
    <tr>
      <th>Item</th>
      <th>Quantity</th>
      <th>Price</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Apple</td>
      <td>10</td>
      <td>2.50</td>
    </tr>
    <tr>
      <td>Banana</td>
      <td>5</td>
      <td>1.20</td>
    </tr>
  </tbody>
</table>

<button onclick="exportToExcel()">Export to Excel</button>

<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.2.0/exceljs.min.js"></script>
<script src="path/to/CustomExport.js"></script>
<script>
async function exportToExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Inventory');

    exportTableToWorksheet('exportTable', worksheet);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'Inventory.xlsx';
    link.click();
}
</script>
```

### Example with hidden columns and row grouping
```html
<table class="exportTable">
    <thead>
        <tr>
            <th class="hide-column">Internal Code</th>
            <th>Item</th>
            <th>Price</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>X123</td>
            <td><b>Apple</b></td>
            <td>1.00</td>
        </tr>
        <tr class="collapse">
            <td>X124</td>
            <td><b>Banana</b></td>
            <td>0.50</td>
        </tr>
    </tbody>
</table>

<button onclick="exportToExcel()">Export to Excel</button>

<script>
  // Setup workbook, wait for table, export
</script>
```

### Example with Merged Cells

```html
<table class="exportTable">
  <thead>
    <tr>
      <th rowspan="2">Product</th>
      <th colspan="2">Q1 Sales</th>
    </tr>
    <tr>
      <th>Volume</th>
      <th>Revenue</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Apple</td>
      <td>120</td>
      <td>$2400</td>
    </tr>
    <tr>
      <td>Banana</td>
      <td>200</td>
      <td>$3000</td>
    </tr>
  </tbody>
</table>

<button onclick="exportToExcel()">Export to Excel</button>

<script>
  // Setup workbook, wait for table, export
</script>
```

## Functions
| Function | Description |
|---------|-------------|
| `delay(ms)` | Returns a promise that resolves after a delay |
| `getColumnAddress(colNum)` | Converts a column number (1 → A, 27 → AA) |
| `columnToNumber(column)` | Converts column letters to numeric index |
| `setCellStyle(htmlCell, excelCell)` | Applies background, font, alignment, borders |
| `getColSpan(worksheet, cellA  ddress, rowNumber)` | Returns the colspan for a merged cell |
| `getMaxRow(worksheet)` | Returns the last used row number |
| `isTopLeftCellOfMerge(worksheet, cellAddress)` | Checks if a cell is the top-left in a merged range |
| `exportTableToWorksheet(tableClassName, worksheet, startRow = 0, excludeNumberFormatting = [])` | Main function to export an HTML table to an Excel sheet |
| `waitForTableLoad(tableId)` | Asynchronously waits for a table to be present in DOM |

## Credits
Thanks to the [ExcelJS](https://github.com/exceljs/exceljs) team for their outstanding open-source library!
