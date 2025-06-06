/*!
 * Custom ExcelJS functions to export Excel with merging
 * 2025 - Devina
 * Version 2.1.0
 */
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function getColumnAddress(colNum) {
    let columnAddress = "";
    let dividend = colNum;
    while (dividend > 0) {
        let remainder = (dividend - 1) % 26;
        columnAddress = String.fromCharCode("A".charCodeAt(0) + remainder) + columnAddress;
        dividend = Math.floor((dividend - remainder) / 26);
    }
    return columnAddress;
}

function setCellStyle(cell, excelCell) {
    let cellHtml = cell.innerHTML.trim();
    let cellText = cellHtml.replace(/&nbsp;/gi, "").trim();

    if (cellText === "" && cellHtml != '&nbsp;') {
        excelCell.value = " ";
        excelCell.fill = undefined;
        excelCell.font = undefined;
        excelCell.alignment = undefined;
        excelCell.border = undefined;
        return;
    }

    let bgColor = getComputedStyle(cell).backgroundColor;

    // if cell bgColor is empty, check parent <tr>
    if (!bgColor || bgColor === "rgba(0, 0, 0, 0)") {
        let row = cell.closest("tr");
        if (row) {
            let rowBgColor = getComputedStyle(row).backgroundColor;
            if (rowBgColor && rowBgColor !== "rgba(0, 0, 0, 0)" && rowBgColor !== "rgb(255, 255, 255)") {
                bgColor = rowBgColor;
            }
        }
    }

    // if neither cell nor row has bgColor, check thead/tbody
    if (!bgColor || bgColor === "rgba(0, 0, 0, 0)") {
        let section = cell.closest("thead, tbody");
        if (section) {
            let sectionBgColor = getComputedStyle(section).backgroundColor;
            if (sectionBgColor && sectionBgColor !== "rgba(0, 0, 0, 0)" && sectionBgColor !== "rgb(255, 255, 255)") {
                bgColor = sectionBgColor;
            }
        }
    }

    let textColor = getComputedStyle(cell).color;
    let transform = window.getComputedStyle(cell).transform;
    let writingMode = window.getComputedStyle(cell).writingMode;
    let alignVertical = "bottom";
    let alignHorizontal = "left";
    let fontWeight = getComputedStyle(cell).fontWeight;
    if (cell.innerHTML.includes("<b") || cell.innerHTML.includes("<strong")) {
        fontWeight = 700;
    }
    let bold = fontWeight >= 700;


    if (bgColor !== "rgba(0, 0, 0, 0)" && bgColor.startsWith("rgb(") && bgColor !== "rgb(255, 255, 255)") {
        let backgroundColor = bgColor.match(/\d+/g).map(Number);
        excelCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF" + backgroundColor.map(c => c.toString(16).padStart(2, "0")).join("") }
        };
    }

    textColor = textColor !== "rgba(0, 0, 0, 0)" && textColor.startsWith("rgb(") && textColor !== "rgb(255, 255, 255)"
        ? "FF" + textColor.match(/\d+/g).map(Number).map(c => c.toString(16).padStart(2, "0")).join("")
        : "FF000000";

    excelCell.font = { name: "Arial", size: 12, bold: bold, color: { argb: textColor } };

    let rotate = transform !== "none" ? getRotationAngle(transform) : 0;
    if (writingMode === "vertical-rl") rotate -= 90;
    if (excelCell.isMerged) alignVertical = alignHorizontal = "middle";
    if (cell.classList.contains("text-center")) alignHorizontal = "center";
    if (cell.classList.contains("text-right")) alignHorizontal = "right";

    excelCell.alignment = { horizontal: alignHorizontal, vertical: alignVertical, textRotation: rotate };
    if (cell.innerHTML.includes("<br")) excelCell.alignment.wrapText = true;
    if (typeof excelCell.value === "string" || typeof excelCell.value === "number") {
        excelCell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
}


function getRotationAngle(transformMatrix) {
    let values = transformMatrix.split("(")[1].split(")")[0].split(",");
    return Math.round(Math.atan2(values[1], values[0]) * (180 / Math.PI));
}

function getColSpan(worksheet, cellAddress, rowNumber) {
    var currentColumn = cellAddress.match(/[A-Z]+/)[0],
        colSpan = 0;
    do {
        colSpan++;
        var nextColumn,
            nextCellAddress = getColumnAddress(columnToNumber(currentColumn) + colSpan) + rowNumber,
            nextCell = worksheet.getCell(nextCellAddress);
    } while (nextCell.isMerged);
    return colSpan;
}

function getMaxRow(worksheet) {
    let maxRow = 0;
    worksheet.eachRow((_, rowNumber) => maxRow = rowNumber);
    return maxRow;
}

function isTopLeftCellOfMerge(worksheet, cellAddress) {
    let cell = worksheet.getCell(cellAddress);
    if (!cell.isMerged) return false;
    let rowNumber = cell.row;
    let columnLetter = cellAddress.match(/[A-Z]+/)[0];

    for (let i = 1; i < columnLetter.length; i++) {
        let prevCell = worksheet.getCell(columnLetter.substring(0, i) + rowNumber);
        if (prevCell.isMerged && cell.master === prevCell.master) return false;
    }
    for (let j = rowNumber - 1; j > 0; j--) {
        let aboveCell = worksheet.getCell(columnLetter + j);
        if (aboveCell.isMerged && cell.master === aboveCell.master) return false;
    }
    return true;
}

function isThRowBelow(cell) {
    let row = cell.parentElement,
        table = row.closest("table"),
        cellIndex = Array.from(row.cells).indexOf(cell),
        allRows = Array.from(table.querySelectorAll("tr"));
    if (cellIndex === row.cells.length - 1) {
        let nextRowIndex = allRows.indexOf(row) + 1;
        if (nextRowIndex < allRows.length) {
            let nextRow = allRows[nextRowIndex];
            if ("th" === cell.tagName.toLowerCase() && nextRow.cells.length > 0 && "th" === nextRow.cells[0].tagName.toLowerCase()) return !0;
        }
    }
    return !1;
}

function columnToNumber(column) {
    return column.split("").reduce((columnNumber, char) =>
        columnNumber * 26 + (char.charCodeAt(0) - ("A".charCodeAt(0) - 1)), 0);
}

function exportTableToWorksheet(tableClassName, worksheet, startRow = 0, excludeNumberFormatting = []) {
    let tables = document.getElementsByClassName(tableClassName);
    let columnAddressCache = {};
    let columnWidths = {};
    let startColumn = 1;

    function getCachedColumnAddress(columnNumber) {
        if (!columnAddressCache[columnNumber]) {
            columnAddressCache[columnNumber] = getColumnAddress(columnNumber);
        }
        return columnAddressCache[columnNumber];
    }

    function getColumnWidth(text) {
        return 8.5 + 3 * Math.log(text.length);
    }


    for (let i = 0; i < tables.length; i++) {
        let table = tables[i];
        var rows = tables[i].rows;
        var maxCols = 0;
        var lastMaxRow = 0;
        var lastCell = null;
        var localMaxCol = 0;
        var hiddenColumns = new Set();

        let headerRows = table.querySelectorAll("thead tr");
        let colTracker = [];

        if (headerRows.length > 0) {
            // Column hiding from header
            for (let r = 0; r < headerRows.length; r++) {
                let row = headerRows[r];
                let cells = Array.from(row.cells);
                let colIndex = 0;

                for (let cell of cells) {
                    // Find the next free column
                    while (colTracker[r]?.[colIndex]) colIndex++;

                    let rowspan = cell.rowSpan || 1;
                    let colspan = cell.colSpan || 1;

                    // Assign Excel column index for this cell
                    for (let i = 0; i < colspan; i++) {
                        for (let j = 0; j < rowspan; j++) {
                            if (!colTracker[r + j]) colTracker[r + j] = [];
                            colTracker[r + j][colIndex + i] = true;
                        }
                    }

                    // If this header cell should trigger column hiding:
                    if (cell.classList.contains("hide-column")) {
                        for (let i = 0; i < colspan; i++) {
                            hiddenColumns.add(colIndex + i + 1);
                        }
                    }

                    colIndex += colspan;
                }
            }
        }

        for (let j = 0; j < rows.length; j++) {
            var adjustedRow = j + startRow + lastMaxRow;
            var row = rows[j];
            var currentCell = startColumn;
            var localMaxCol = 0;
            var excelRow = worksheet.getRow(adjustedRow + 1);
            var isLastCell = false;

            // Set grouping level
            if (row.classList.contains('collapse')) {
                excelRow.outlineLevel = 1;
                if (row.classList.contains('in')) {
                    excelRow.hidden = false;
                } else {
                    excelRow.hidden = true;
                }
            }

            if (row.classList.contains("hidden")) {
                excelRow.hidden = true;
            }

            for (let k = 0; k < row.cells.length; k++) {
                var cell = row.cells[k];
                let cellValue;
                if (cell.innerHTML.includes("<br")) {
                    const temp = document.createElement("div");
                    temp.innerHTML = cell.innerHTML.replace(/<br\s*\/?>/gi, "\n");
                    cellValue = temp.textContent || temp.innerText || "";
                } else {
                    cellValue = cell.textContent.trim();
                }

                // find the next unmerged cell
                var nextAvailableColumn = currentCell;
                while (worksheet.getCell(`${getColumnAddress(nextAvailableColumn)}${adjustedRow + 1}`).isMerged) {
                    nextAvailableColumn++; // skip merged
                }
                var cellAddress = `${getColumnAddress(nextAvailableColumn)}${adjustedRow + 1}`;
                var excelCell = worksheet.getCell(cellAddress);

                var rowSpan = cell.rowSpan || 1;
                var colSpan = cell.colSpan || 1;

                if (rowSpan > 1 || colSpan > 1) {
                    let mergeStart = cellAddress;
                    let mergeEnd = getColumnAddress(nextAvailableColumn + colSpan - 1) + (adjustedRow + rowSpan);
                    worksheet.mergeCells(mergeStart, mergeEnd);
                }

                // check if column is excluded from dynamic number formatting
                if (!excludeNumberFormatting.includes(getColumnAddress(nextAvailableColumn))) {
                    // format numbers and percentage
                    if (cellValue.trim().replace(/,/g, "").endsWith("%")) {
                        cellValue = parseFloat(cellValue.trim().replace(/,/g, "").replace("%", "")) / 100;
                        excelCell.numFmt = Number.isInteger(100 * cellValue) ? "0%" : "0.00%";
                    } else if (!isNaN(cellValue.trim().replace(/,/g, "")) && cellValue.trim() !== "") {
                        cellValue = parseFloat(cellValue.trim().replace(/,/g, ""));
                        excelCell.numFmt = Number.isInteger(cellValue) ? "#,##0" : "#,##0.00";
                    }
                }

                // set dynamic column widths
                if (cellValue !== " " && (!excelCell.isMerged || isTopLeftCellOfMerge(worksheet, cellAddress))) {
                    let columnWidth = getColumnWidth(cellValue.toString().replace(/,/g, ""));
                    if (!columnWidths[currentCell] || columnWidth >= columnWidths[currentCell]) {
                        columnWidths[currentCell] = columnWidth;
                    }
                } else if (!columnWidths[currentCell] || columnWidths[currentCell] <= 8.3) {
                    columnWidths[currentCell] = 8.3;
                }

                excelCell.value = typeof cellValue === "string" ? cellValue.trim() : cellValue;
                setCellStyle(cell, excelCell);

                currentCell = nextAvailableColumn + colSpan; // Move to the next available column
            }

            (maxCols = Math.max(maxCols, localMaxCol)), isLastCell ? ((lastMaxRow = 0), (lastCell = cell)) : (lastCell = null);

        }

        excelRow.commit(), (startColumn = localMaxCol + 2);
    }

    // apply column widths
    for (let columnIndex in columnWidths) {
        let columnLetter = getColumnAddress(parseInt(columnIndex));
        worksheet.getColumn(columnLetter).width = columnWidths[columnIndex] + 2;
    }

    for (let col of hiddenColumns) {
        worksheet.getColumn(col).hidden = true;
    }

    worksheet.views = [{ zoomScale: 70 }];
}

async function waitForTableLoad(tableId) {
    return (
        await delay(500),
        new Promise((resolve, reject) => {
            let attempts = 0;
            function checkTable() {
                let tables = document.querySelectorAll("table");
                let foundTable = null;
                tables.forEach((table) => {
                    Array.from(table.classList).some((className) => className.startsWith(tableId)) && (foundTable = table);
                }),
                    foundTable && foundTable.rows.length > 0 ? resolve() : attempts < 50 ? (attempts++, setTimeout(checkTable, 1e3)) : reject(new Error("Table did not load in time."));
            }
            checkTable();
        })
    );
}
