import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

/**
 * Generic utility for exporting data to Excel with advanced features
 */

// Helper: Apply border to cell
const applyBorder = (cell, color, style = 'thin') => {
    cell.border = {
        top: { style, color: { argb: color } },
        left: { style, color: { argb: color } },
        bottom: { style, color: { argb: color } },
        right: { style, color: { argb: color } },
    };
};

// Helper: Apply style to cell
const applyCellStyle = (cell, style) => {
    if (style.font) cell.font = { ...cell.font, ...style.font };
    if (style.fill) cell.fill = style.fill;
    if (style.alignment) cell.alignment = style.alignment;
    if (style.border) cell.border = style.border;
    if (style.numFmt) cell.numFmt = style.numFmt;
};

// Helper: Apply style to row
const applyRowStyle = (row, style) => {
    if (style.font) row.font = style.font;
    if (style.alignment) row.alignment = style.alignment;
    if (style.height) row.height = style.height;

    row.eachCell((cell) => {
        if (style.fill) cell.fill = style.fill;
        if (style.border) cell.border = style.border;
    });
};

// Helper: Parse customData format
const parseCustomData = (customData) => {
    const dataArray = Array.isArray(customData) ? customData : (customData?.data || []);
    const style = !Array.isArray(customData) ? customData?.style : null;
    return { dataArray, style };
};

// Helper: Check if cell is in range
const isCellInRange = (cellAddr, startCell, endCell) => {
    const getCellCoords = (addr) => {
        if (!addr || typeof addr !== 'string') return null;
        const match = addr.match(/^([A-Z]+)(\d+)$/);
        if (!match) return null;
        const col = match[1].split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
        const row = parseInt(match[2]);
        if (isNaN(row) || row < 1) return null;
        return { col, row };
    };

    const cell = getCellCoords(cellAddr);
    const start = getCellCoords(startCell);
    const end = getCellCoords(endCell);

    if (!cell || !start || !end) return false;

    return cell.row >= start.row && cell.row <= end.row &&
        cell.col >= start.col && cell.col <= end.col;
};

const exportToExcel = async (config) => {
    const { fileName = 'export.xlsx', sheets = [] } = config;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'ExcelJS App';
    workbook.created = new Date();

    // Process each sheet independently - no cross-sheet dependencies
    sheets.forEach((sheetConfig) => {
        // Extract only this sheet's configuration (independent from other sheets)
        const {
            sheetName = 'Sheet1',
            data = [],
            customData = [],
            columns = [],
            headerStyle = {},
            dataStyle = null,
            freeze = null,
            dropdowns = [],
            comments = [],
            lockedColumns = [],
            columnColors = {},
            rowStyles = [],
            cellStyles = [],
            protectSheet = null,
        } = sheetConfig;

        // Create worksheet for this sheet only
        const worksheet = workbook.addWorksheet(sheetName);

        // Parse custom data once
        const { dataArray: customDataArray, style: customDataStyle } = parseCustomData(customData);

        // Only use provided columns - DO NOT auto-generate (independent behavior)
        const hasColumns = columns.length > 0;
        const finalColumns = hasColumns ? columns : [];

        // Create column index map for O(1) lookups (independent)
        const columnMap = new Map(finalColumns.map((col, idx) => [col.key, idx]));

        // Define columns ONLY if explicitly provided
        if (hasColumns) {
            worksheet.columns = finalColumns.map(({ header, key, width = 15 }) => ({ header, key, width }));
        }

        // Add main data rows (independent - works with or without columns)
        const dataStartRow = hasColumns ? 2 : 1;
        data.forEach((row, index) => {
            if (hasColumns) {
                worksheet.addRow(row);
            } else {
                // Add raw data without column mapping
                // Note: Order depends on object key iteration order
                const values = Object.values(row);
                if (values.length === 0 && index === 0) {
                    console.warn('First data row is empty - this may cause inconsistent column layout');
                }
                worksheet.addRow(values);
            }
        });
        const dataEndRow = data.length > 0 ? dataStartRow + data.length - 1 : dataStartRow;

        // Apply style to all data rows (independent)
        if (dataStyle && data.length > 0) {
            for (let rowNum = dataStartRow; rowNum <= dataEndRow; rowNum++) {
                const targetRow = worksheet.getRow(rowNum);
                applyRowStyle(targetRow, dataStyle);
            }
        }

        // Add custom data rows with optional styling (independent)
        customDataArray.forEach((row) => {
            let addedRow;
            if (hasColumns) {
                addedRow = worksheet.addRow(row);
            } else {
                // Add raw data without column mapping
                const values = Object.values(row);
                addedRow = worksheet.addRow(values);
            }

            if (customDataStyle && addedRow) {
                Object.assign(addedRow, {
                    font: customDataStyle.font || {},
                    alignment: customDataStyle.alignment || {},
                    height: customDataStyle.height,
                });

                addedRow.eachCell((cell) => {
                    if (customDataStyle.fill) cell.fill = customDataStyle.fill;
                    if (customDataStyle.border) cell.border = customDataStyle.border;
                });
            }
        });

        // Style header row (independent - only if columns explicitly provided)
        if (hasColumns && finalColumns.length > 0) {
            const headerRow = worksheet.getRow(1);

            // Build default header font
            const defaultFont = {
                bold: true,
                size: 12,
                color: { argb: headerStyle.textColor || 'FFFFFFFF' }
            };

            // Merge with custom font (custom takes precedence)
            const finalFont = headerStyle.font ? { ...defaultFont, ...headerStyle.font } : defaultFont;

            // Apply row-level styles
            Object.assign(headerRow, {
                font: finalFont,
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: headerStyle.backgroundColor || 'FF4472C4' } },
                alignment: headerStyle.alignment || { vertical: 'middle', horizontal: 'center' },
                height: headerStyle.height,
            });

            // Apply border to each cell
            headerRow.eachCell((cell) => {
                // Apply custom border if provided, otherwise use borderColor and borderStyle
                if (headerStyle.border) {
                    cell.border = headerStyle.border;
                } else {
                    const borderColor = headerStyle.borderColor || 'FF000000';
                    const borderStyle = headerStyle.borderStyle || 'thin'; // thin, medium, thick
                    // Add FF prefix if not present (handle both 'FF000000' and '000000')
                    const argbColor = borderColor.length === 6 ? 'FF' + borderColor : borderColor;
                    applyBorder(cell, argbColor, borderStyle);
                }
            });

            headerRow.commit();
        }

        // Add comments to header (independent - only if columns provided)
        if (comments.length > 0 && hasColumns) {
            comments.forEach(({ key, text }) => {
                const columnIndex = columnMap.get(key);
                if (columnIndex !== undefined) {
                    worksheet.getCell(1, columnIndex + 1).note = {
                        texts: [{ font: { size: 12, color: { theme: 1 }, name: 'Calibri', family: 2 }, text }],
                        margins: { insetmode: 'auto', inset: [0.13, 0.13, 0.13, 0.13] },
                        protection: { locked: 'True', lockText: 'False' },
                        editAs: 'absolute',
                    };
                }
            });
        }

        // Add dropdowns (independent - only if columns provided)
        if (dropdowns.length > 0 && hasColumns) {
            const totalDataRows = data.length + customDataArray.length;

            dropdowns.forEach(({ key, options, startRow = 2, endRow, additionalRows = 100 }) => {
                const columnIndex = columnMap.get(key);
                if (columnIndex === undefined) return;
                if (!options || options.length === 0) return; // Edge case: empty options

                const finalEndRow = endRow || (totalDataRows + 1 + additionalRows);

                // Escape quotes and handle special characters
                const escapedOptions = options.map(opt => String(opt).replace(/"/g, '""'));
                const formulaString = `"${escapedOptions.join(',')}"`;

                // Excel has ~255 char limit for formula strings
                if (formulaString.length > 255) {
                    console.warn(`Dropdown for column "${key}" has too many options (${formulaString.length} chars). Consider using a named range instead.`);
                }

                const validation = {
                    type: 'list',
                    allowBlank: true,
                    formulae: [formulaString],
                    showErrorMessage: true,
                    errorStyle: 'error',
                    errorTitle: 'Invalid Selection',
                    error: 'Please select a value from the dropdown',
                };

                for (let row = startRow; row <= finalEndRow; row++) {
                    worksheet.getCell(row, columnIndex + 1).dataValidation = validation;
                }
            });
        }

        // Apply column colors (independent - skip header row, only if columns provided)
        if (Object.keys(columnColors).length > 0 && hasColumns) {
            Object.entries(columnColors).forEach(([columnKey, { backgroundColor, textColor, numFmt }]) => {
                const columnIndex = columnMap.get(columnKey);
                if (columnIndex === undefined) return;

                const fill = backgroundColor ? { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } } : null;

                for (let row = 2; row <= worksheet.rowCount; row++) {
                    const cell = worksheet.getCell(row, columnIndex + 1);
                    if (fill) cell.fill = fill;
                    if (textColor) cell.font = { ...cell.font, color: { argb: textColor } };
                    if (numFmt) cell.numFmt = numFmt;
                }
            });
        }

        // Apply row styles (independent)
        if (rowStyles.length > 0) {
            rowStyles.forEach((rowStyleConfig) => {
                const { row, rows, startRow, endRow, style } = rowStyleConfig;

                if (row !== undefined) {
                    // Single row
                    const targetRow = worksheet.getRow(row);
                    applyRowStyle(targetRow, style);
                } else if (rows && Array.isArray(rows)) {
                    // Multiple specific rows
                    rows.forEach((rowNum) => {
                        const targetRow = worksheet.getRow(rowNum);
                        applyRowStyle(targetRow, style);
                    });
                } else if (startRow !== undefined && endRow !== undefined) {
                    // Range of rows
                    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
                        const targetRow = worksheet.getRow(rowNum);
                        applyRowStyle(targetRow, style);
                    }
                }
            });
        }

        // Apply cell styles (independent)
        if (cellStyles.length > 0) {
            cellStyles.forEach((cellStyleConfig) => {
                const { cell, cells, row, column, startCell, endCell, style } = cellStyleConfig;

                if (cell) {
                    // Single cell by address (e.g., 'A1')
                    const targetCell = worksheet.getCell(cell);
                    applyCellStyle(targetCell, style);
                } else if (cells && Array.isArray(cells)) {
                    // Multiple specific cells
                    cells.forEach((cellAddr) => {
                        const targetCell = worksheet.getCell(cellAddr);
                        applyCellStyle(targetCell, style);
                    });
                } else if (row !== undefined && column !== undefined) {
                    // Cell by row number and column key (only works when columns are defined)
                    if (hasColumns) {
                        const columnIndex = columnMap.get(column);
                        if (columnIndex !== undefined) {
                            const targetCell = worksheet.getCell(row, columnIndex + 1);
                            applyCellStyle(targetCell, style);
                        }
                    }
                } else if (startCell && endCell) {
                    // Range of cells (e.g., 'A1:C3')
                    const range = worksheet.getCell(startCell).address + ':' + worksheet.getCell(endCell).address;
                    worksheet.eachRow((row, rowNumber) => {
                        row.eachCell((cell) => {
                            const cellAddr = cell.address;
                            if (isCellInRange(cellAddr, startCell, endCell)) {
                                applyCellStyle(cell, style);
                            }
                        });
                    });
                }
            });
        }

        // Freeze panes (independent)
        if (freeze?.rows > 0 || freeze?.columns > 0) {
            worksheet.views = [{
                state: 'frozen',
                xSplit: freeze.columns || 0,
                ySplit: freeze.rows || 0,
                activeCell: 'A1'
            }];
        }

        // Protect entire sheet (independent)
        if (protectSheet) {
            const protectionOptions = typeof protectSheet === 'object' ? {
                password: protectSheet.password || '',
                selectLockedCells: protectSheet.selectLockedCells !== false,
                selectUnlockedCells: protectSheet.selectUnlockedCells !== false,
                formatCells: protectSheet.formatCells || false,
                formatColumns: protectSheet.formatColumns || false,
                formatRows: protectSheet.formatRows || false,
                insertColumns: protectSheet.insertColumns || false,
                insertRows: protectSheet.insertRows || false,
                deleteColumns: protectSheet.deleteColumns || false,
                deleteRows: protectSheet.deleteRows || false,
                sort: protectSheet.sort || false,
                autoFilter: protectSheet.autoFilter || false,
            } : {
                password: '',
                selectLockedCells: true,
                selectUnlockedCells: true,
                formatCells: false,
                formatColumns: false,
                formatRows: false,
                insertColumns: false,
                insertRows: false,
                deleteColumns: false,
                deleteRows: false,
            };

            worksheet.protect(protectionOptions.password, protectionOptions);
        }

        // Lock specific columns (independent - only if columns provided)
        if (lockedColumns.length > 0 && hasColumns) {
            // If sheet not already protected, protect it now
            if (!protectSheet) {
                worksheet.protect('', {
                    selectLockedCells: true,
                    selectUnlockedCells: true,
                    formatCells: false,
                    formatColumns: false,
                    formatRows: false,
                    insertColumns: false,
                    insertRows: false,
                    deleteColumns: false,
                    deleteRows: false,
                });
            }

            // Unlock all cells first
            worksheet.eachRow((row) => row.eachCell((cell) => cell.protection = { locked: false }));

            // Lock specific columns
            lockedColumns.forEach((columnKey) => {
                const columnIndex = columnMap.get(columnKey);
                if (columnIndex !== undefined) {
                    for (let row = 1; row <= worksheet.rowCount; row++) {
                        worksheet.getCell(row, columnIndex + 1).protection = { locked: true };
                    }
                }
            });
        }
    });

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, fileName);
};

export default exportToExcel;
