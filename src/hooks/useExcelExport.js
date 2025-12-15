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

// Helper: Parse customData format
const parseCustomData = (customData) => {
    const dataArray = Array.isArray(customData) ? customData : (customData?.data || []);
    const style = !Array.isArray(customData) ? customData?.style : null;
    return { dataArray, style };
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
            freeze = null,
            dropdowns = [],
            comments = [],
            lockedColumns = [],
            columnColors = {},
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
        data.forEach((row) => {
            if (hasColumns) {
                worksheet.addRow(row);
            } else {
                // Add raw data without column mapping
                const values = Object.values(row);
                worksheet.addRow(values);
            }
        });

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
            Object.assign(headerRow, {
                font: { bold: true, size: 12, color: { argb: headerStyle.textColor || 'FFFFFFFF' }, ...headerStyle.font },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: headerStyle.backgroundColor || 'FF4472C4' } },
                alignment: { vertical: 'middle', horizontal: 'center', ...headerStyle.alignment },
            });

            headerRow.eachCell((cell) => applyBorder(cell, headerStyle.borderColor || 'FF000000'));
            if (headerStyle.border) headerRow.border = { ...headerRow.border, ...headerStyle.border };
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

                const finalEndRow = endRow || (totalDataRows + 1 + additionalRows);
                const validation = {
                    type: 'list',
                    allowBlank: true,
                    formulae: [`"${options.join(',')}"`],
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
            Object.entries(columnColors).forEach(([columnKey, { backgroundColor, textColor }]) => {
                const columnIndex = columnMap.get(columnKey);
                if (columnIndex === undefined) return;

                const fill = backgroundColor ? { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } } : null;

                for (let row = 2; row <= worksheet.rowCount; row++) {
                    const cell = worksheet.getCell(row, columnIndex + 1);
                    if (fill) cell.fill = fill;
                    if (textColor) cell.font = { ...cell.font, color: { argb: textColor } };
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

        // Protect sheet and lock specific columns (independent - only if columns provided)
        if (lockedColumns.length > 0 && hasColumns) {
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
