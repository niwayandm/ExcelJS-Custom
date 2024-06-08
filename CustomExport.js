/*!
* Custom ExcelJS functions to export Excel with merging.
* 2024 - Devina
*/

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Function for handling columns greater than z
function getColumnAddress(colNum) {
    let columnAddress = '';
    let dividend = colNum;
    let remainder;

    while (dividend > 0) {
        remainder = (dividend - 1) % 26;
        columnAddress = String.fromCharCode('A'.charCodeAt(0) + remainder) + columnAddress;
        dividend = parseInt((dividend - remainder) / 26);
    }

    return columnAddress;
}

function setCellStyle(htmlCell, excelCell) {
    // Set cell background color
    var bgColor = getComputedStyle(htmlCell).backgroundColor;
    if (bgColor !== 'rgba(0, 0, 0, 0)' && bgColor.startsWith('rgb(') && bgColor !== 'rgb(255, 255, 255)') {
        var color = bgColor.match(/\d+/g).map(Number);
        excelCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'FF' + color.map(c => c.toString(16).padStart(2, '0')).join('')
            }
        };

        // Set font
        excelCell.font = {
            name: 'Arial',
            size: 10,
            bold: true
        };
    }

    // Set alignment
    if (excelCell.isMerged) {
        excelCell.alignment = {
            vertical: 'middle'
        };
    }

    if (htmlCell.classList.contains('text-center')) {
        excelCell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
    } else if (htmlCell.classList.contains('text-right')) {
        excelCell.alignment = {
            horizontal: 'right',
        };
    }

    // Set cell borders
    if (typeof excelCell.value === 'string') {
        var trimmedValue = excelCell.value.replace(/\u00a0/g, '').trim();
        if (trimmedValue !== '' || htmlCell.innerHTML.trim() === '&nbsp;') {
            excelCell.border = {
                top: {
                    style: 'thin'
                },
                left: {
                    style: 'thin'
                },
                bottom: {
                    style: 'thin'
                },
                right: {
                    style: 'thin'
                }
            };
        }
    } else if (typeof excelCell.value === 'number') {
        excelCell.border = {
            top: {
                style: 'thin'
            },
            left: {
                style: 'thin'
            },
            bottom: {
                style: 'thin'
            },
            right: {
                style: 'thin'
            }
        };
    }

}

function getColSpan(worksheet, cellAddress, rowNumber) {
    var currentColumn = cellAddress.match(/[A-Z]+/)[0];
    var colSpan = 0;

    do {
        colSpan++;
        var nextColumn = getColumnAddress(columnToNumber(currentColumn) + colSpan);
        var nextCellAddress = nextColumn + rowNumber;
        var nextCell = worksheet.getCell(nextCellAddress);
    } while (nextCell.isMerged);

    return colSpan;
}

function getMaxRow(worksheet) {
    var maxRow = 0;
    worksheet.eachRow(function (row, rowNumber) {
        maxRow = rowNumber;
    });
    return maxRow;
}

function isTopLeftCellOfMerge(worksheet, cellAddress) {
    var cell = worksheet.getCell(cellAddress);
    if (!cell.isMerged) return false;

    var rowNumber = cell.row;
    var columnLetter = cellAddress.match(/[A-Z]+/)[0];

    // Check if any previous cell in the same row is merged and shares the merge area
    for (var i = 1; i < columnLetter.length; i++) {
        var prevCellAddress = columnLetter.substring(0, i) + rowNumber;
        var prevCell = worksheet.getCell(prevCellAddress);
        if (prevCell.isMerged && cell.master === prevCell.master) {
            return false;
        }
    }

    // Check if any previous cell in the same column is merged and shares the merge area
    for (var j = rowNumber - 1; j > 0; j--) {
        var aboveCellAddress = columnLetter + j;
        var aboveCell = worksheet.getCell(aboveCellAddress);
        if (aboveCell.isMerged && cell.master === aboveCell.master) {
            return false;
        }
    }

    return true;
}

function columnToNumber(column) {
    let columnNumber = 0;
    for (let i = 0; i < column.length; i++) {
        columnNumber *= 26;
        columnNumber += (column.charCodeAt(i) - ('A'.charCodeAt(0) - 1));
    }
    return columnNumber;
}

function extractTextFromHtml(html) {
    // Create a temporary div element to parse the HTML content
    var tempDiv = document.createElement("div");
    tempDiv.innerHTML = html.replace(/&nbsp;/g, '\u00A0'); // Convert &nbsp; to Unicode

    var text = tempDiv.textContent || tempDiv.innerText || "";
    return text;
}

function exportTableToWorksheet2(tableClassName, worksheet, startRow = 0) {
    var tables = document.getElementsByClassName(tableClassName);
    var columnAddressCache = {};
    var startColumn = 1; // Global tracker for the starting column of each new table
    var columnWidths = {}; // Object to track max widths of columns

    function getCachedColumnAddress(columnNumber) {
        if (!columnAddressCache[columnNumber]) {
            columnAddressCache[columnNumber] = getColumnAddress(columnNumber);
        }
        return columnAddressCache[columnNumber];
    }

    function getColumnWidth(text) {
        return text.length * 1.75;
    }

    for (var i = 0; i < tables.length; i++) {
        var table = tables[i];
        var rows = table.rows;
        var maxCols = 0;
        var maxRows = 0;


        for (var j = 0; j < rows.length; j++) {
            var adjustedRow = j + startRow + maxRows;
            var row = rows[j];
            var currentCell = startColumn;
            var cells = row.cells;
            var localMaxCol = 0;

            var excelRow = worksheet.getRow(adjustedRow + 1);

            // Set grouping level
            if (row.classList.contains('collapse') && !row.classList.contains('SU') && !row.classList.contains('PF')) {
                excelRow.outlineLevel = 1;
                if (row.classList.contains('in')) {
                    excelRow.hidden = false;
                } else {
                    excelRow.hidden = true;
                }
            } else if (row.classList.contains('PF')) {
                excelRow.hidden = true;
            }

            for (var k = 0; k < cells.length; k++) {
                var cell = cells[k];
                var cellValue = cell.innerText || cell.textContent;
                var cellAddress = `${getColumnAddress(currentCell)}${adjustedRow + 1}`;
                var excelCell = worksheet.getCell(cellAddress);
                var isTopLeft = isTopLeftCellOfMerge(worksheet, cellAddress);

                // console.log("Initial address: " + cellAddress)

                // Skip processing if the cell is merged and not the top-left cell of the merge area
                if (excelCell.isMerged && !isTopLeft) {
                    // console.log("Skipping: " + cellAddress + " as it's not the top-left of the merge.");

                    var colSpan = getColSpan(worksheet, cellAddress, adjustedRow + 1)
                    adjustedRow -= maxRows;
                    currentCell += colSpan;

                    maxRows = 0;

                    cellAddress = `${getColumnAddress(currentCell)}${adjustedRow + 1}`;
                    excelCell = worksheet.getCell(cellAddress);

                    // console.log("New Address: " + cellAddress);

                }

                // Handle merged cells
                var rowSpan = cell.rowSpan || 1;
                var colSpan = cell.colSpan || 1;
                if (rowSpan > 1 || colSpan > 1) {
                    var mergeStart = cellAddress;
                    var mergeEnd = getCachedColumnAddress(currentCell + colSpan - 1) + (adjustedRow + rowSpan);

                    // console.log("Merging: " + mergeStart + " to " + mergeEnd);
                    worksheet.mergeCells(mergeStart, mergeEnd);
                }

                // Change cell type to numeric or percentage

                if (cellValue.trim().replace(/,/g, '').endsWith('%') && !isNaN(cellValue.trim().replace(/,/g, '').replace('%', '')) && cellValue.trim().replace(/,/g, '').replace('%', '').length > 1) {
                    cellValue = parseFloat(cellValue.trim().replace(/,/g, '').replace('%', '')) / 100;
                    if (Number.isInteger(cellValue * 100)) {
                        excelCell.numFmt = '0%';
                    } else {
                        excelCell.numFmt = '0.00%';
                    }
                } else if (!isNaN(cellValue.trim().replace(/,/g, '')) && cellValue.trim() !== '') {
                    cellValue = cellValue.trim().replace(/,/g, '');
                    cellValue = parseFloat(cellValue);

                    if (Number.isInteger(cellValue)) {
                        excelCell.numFmt = '#,##0';
                    } else {
                        excelCell.numFmt = '#,##0.00';
                    }

                }

                excelCell.value = cellValue;


                // Handle cell styling
                setCellStyle(cell, excelCell);

                // Set or update the max width for the current column
                if ((rowSpan > 1 || colSpan > 1) && (cell.tagName.toLowerCase() !== 'th')) {

                } else if (excelCell.value !== '' || (!excelCell.isMerged && !isTopLeft)) {
                    if (tableClassName.includes("2") && tableClassName.includes("22")) {

                    } else {
                        var columnWidth = getColumnWidth(cellValue.toString().replace(/,/g, ''));
                        if (!columnWidths[currentCell] || columnWidth >= columnWidths[currentCell]) {
                            columnWidths[currentCell] = columnWidth;
                        }
                    }

                } else {
                    if (!columnWidths[currentCell] || 8.3 >= columnWidths[currentCell]) {
                        columnWidths[currentCell] = 8.3;
                    }
                }

                currentCell += colSpan;
                localMaxCol = Math.max(localMaxCol, currentCell - 1);
            }
            maxCols = Math.max(maxCols, localMaxCol);

            if (String(excelCell.value).includes("TOTAL")) {
                // console.log(rowSpan - 1);
                maxRows = 0;
            } else {
                maxRows += (rowSpan - maxRows > 0) ? rowSpan - 1 : 0;
            }

        }

        excelRow.commit();
        startColumn = localMaxCol + 2;
    }

    // Set column widths by max widths for each column
    for (let columnIndex in columnWidths) {
        var columnLetter = getColumnAddress(parseInt(columnIndex));
        var column = worksheet.getColumn(columnLetter);
        column.width = columnWidths[columnIndex] + 2;

        // console.log(`${columnLetter} has width of ${columnWidths[columnIndex] + 2}`)
    }

    worksheet.views = [{
        zoomScale: 70
    }];
}

async function waitForTableLoad(tableId) {
    await delay(500);

    return new Promise((resolve, reject) => {
        let attempts = 0;

        function checkTable() {
            const table = document.querySelector(tableId);
            console.log(table)
            if (table && table.rows.length > 0) {
                resolve();
            } else if (attempts < 50) {
                attempts++;
                setTimeout(checkTable, 1000);
            } else {
                reject(new Error('Table did not load in time.'));
            }
        }

        checkTable();
    });
}
