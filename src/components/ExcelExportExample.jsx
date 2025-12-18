import React from 'react';
import exportToExcel from '../hooks/useExcelExport';

const ExcelExportExample = () => {

    const buttonBaseStyle = {
        padding: '12px 20px',
        color: 'white',
        border: 'none',
        borderRadius: '5px',
        cursor: 'pointer',
        fontSize: '16px',
    };

    const sampleData = [
        { id: 1, name: 'John Doe', department: 'Sales', status: 'Active', salary: 50000 },
        { id: 2, name: 'Jane Smith', department: 'IT', status: 'Active', salary: 65000 },
        { id: 3, name: 'Bob Johnson', department: 'HR', status: 'Inactive', salary: 55000 },
        { id: 4, name: 'Alice Brown', department: 'Sales', status: 'Active', salary: 52000 },
        { id: 5, name: 'Charlie Wilson', department: 'IT', status: 'Active', salary: 70000 },
    ];

    const handleExportBasic = async () => {
        const config = {
            fileName: 'basic_export.xlsx',
            sheets: [
                {
                    sheetName: 'Employees',
                    columns: [
                        { header: 'ID', key: 'id', width: 10 },
                        { header: 'Name', key: 'name', width: 20 },
                        { header: 'Department', key: 'department', width: 15 },
                        { header: 'Status', key: 'status', width: 15 },
                        { header: 'Salary', key: 'salary', width: 15 },
                    ],
                    data: sampleData,
                },
            ],
        };

        await exportToExcel(config);
    };

    const handleExportWithAllFeatures = async () => {
        const config = {
            fileName: 'advanced_export.xlsx',
            sheets: [
                {
                    sheetName: 'Employees',
                    columns: [
                        { header: 'ID', key: 'id', width: 10 },
                        { header: 'Name', key: 'name', width: 20 },
                        { header: 'Department', key: 'department', width: 15 },
                        { header: 'Status', key: 'status', width: 15 },
                        { header: 'Salary', key: 'salary', width: 15 },
                    ],
                    data: sampleData,

                    // Style for all data rows (excludes header)
                    dataStyle: {
                        border: {
                            top: { style: 'thin', color: { argb: '000000' } },
                            left: { style: 'thin', color: { argb: '000000' } },
                            bottom: { style: 'thin', color: { argb: '000000' } },
                            right: { style: 'thin', color: { argb: '000000' } },
                        },
                        alignment: {
                            vertical: 'middle',
                            horizontal: 'left',
                        },
                    },

                    // Custom data - additional rows appended below main data
                    // Can be array or object with styling
                    customData: {
                        data: [
                            {},
                            { id: 6, name: 'David Lee', department: 'Finance', status: 'Active', salary: 60000 },
                            { id: 7, name: 'Emma Davis', department: 'Sales', status: 'On Leave', salary: 48000 },
                        ],
                        style: {
                            font: {
                                bold: true,
                                italic: true,
                                color: { argb: 'FF0066CC' },
                            },
                            fill: {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFE8F4F8' }, // Light blue background
                            },
                            height: 20,
                            alignment: {
                                vertical: 'middle',
                                horizontal: 'left',
                            },
                            border: {
                                top: { style: 'thin', color: { argb: 'FF0066CC' } },
                                bottom: { style: 'thin', color: { argb: 'FF0066CC' } },
                            },
                        },
                    },

                    // Custom header styling
                    headerStyle: {
                        backgroundColor: 'FF2E75B6', // Blue background
                        textColor: 'FFFFFFFF', // White text
                        borderColor: 'FF000000', // Black border (or just '000000')
                        borderStyle: 'medium', // thin, medium, thick, double
                        font: {
                            bold: true,
                            size: 13,
                        },
                        alignment: {
                            vertical: 'middle',
                            horizontal: 'left',
                        },
                    },

                    // Comments on header cells
                    comments: [
                        { key: 'id', text: 'Employee ID - Auto generated', author: 'System' },
                        { key: 'department', text: 'Department: Sales, IT, HR, Finance', author: 'Admin' },
                        { key: 'salary', text: 'Annual salary in USD', author: 'HR' },
                    ],

                    // Dropdowns for specific columns
                    dropdowns: [
                        {
                            key: 'department', // Department column
                            options: ['Sales', 'IT', 'HR', 'Finance'],
                            startRow: 4,      // Start from row 4 instead of 2
                            endRow: 20,       // Only up to row 20
                            // startRow and endRow are optional - will auto-calculate based on data
                            // additionalRows: 50 (optional) - adds extra rows below data for new entries
                        },
                        {
                            key: 'status', // Status column
                            options: ['Active', 'Inactive', 'On Leave'],
                        },
                    ],

                    // Freeze first row (header)
                    freeze: {
                        rows: 1,
                        columns: 0,
                    },

                    // // Protect entire sheet (locks all cells by default)
                    // protectSheet: {
                    //     password: 'secret123',
                    //     selectLockedCells: true,
                    //     selectUnlockedCells: true,
                    // },

                    // Lock specific columns (ID column) - requires sheet protection
                    lockedColumns: ['id', 'name'],

                    // Apply colors to specific columns
                    columnColors: {
                        status: {
                            backgroundColor: 'FFFFEAA7', // Light yellow
                            textColor: 'FF000000', // Black text
                        },
                        salary: {
                            backgroundColor: 'FFFFEAA7', // Light yellow
                            textColor: 'FF000000', // Black text
                            numFmt: '$#,##0.00', // Currency format
                        },
                    },

                    // Apply styles to specific rows
                    rowStyles: [
                        {
                            row: 3, // Style row 3
                            style: {
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } },
                                font: { bold: true },
                            },
                        },
                        {
                            startRow: 5,
                            endRow: 6,
                            style: {
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } },
                            },
                        },
                    ],

                    // Apply styles to specific cells
                    cellStyles: [
                        {
                            row: 2,
                            column: 'name',
                            style: {
                                font: { bold: true, color: { argb: 'FFFF0000' } },
                            },
                        },
                        {
                            cell: 'A4',
                            style: {
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
                            },
                        },
                    ],
                },
                {
                    sheetName: 'Summary',
                    columns: [
                        { header: 'Department', key: 'dept', width: 20 },
                        { header: 'Employee Count', key: 'count', width: 20 },
                        { header: 'Total Salary', key: 'totalSalary', width: 20 },
                    ],
                    data: [
                        { dept: 'Sales', count: 2, totalSalary: 102000 },
                        { dept: 'IT', count: 2, totalSalary: 135000 },
                        { dept: 'HR', count: 1, totalSalary: 55000 },
                    ],
                    headerStyle: {
                        backgroundColor: 'FF00B894', // Green background
                        textColor: 'FFFFFFFF',
                    },
                    freeze: {
                        rows: 1,
                        columns: 1,
                    },
                },
                {
                    sheetName: 'Summary1',
                    // Protect entire sheet (locks all cells by default)
                    protectSheet: {
                        password: 'secret123',
                        selectLockedCells: true,
                        selectUnlockedCells: true,
                    },
                    customData: {
                        data: [
                            {},
                            { id: 6, name: 'David Lee', department: 'Finance', status: 'Active', salary: 60000 },
                            { id: 7, name: 'Emma Davis', department: 'Sales', status: 'On Leave', salary: 48000 },
                        ],
                        style: {
                            font: {
                                bold: true,
                                italic: true,
                                color: { argb: 'FF0066CC' },
                            },
                            fill: {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFE8F4F8' }, // Light blue background
                            },
                            height: 20,
                            alignment: {
                                vertical: 'middle',
                                horizontal: 'left',
                            },
                            border: {
                                top: { style: 'thin', color: { argb: 'FF0066CC' } },
                                bottom: { style: 'thin', color: { argb: 'FF0066CC' } },
                            },
                        },
                    },
                },
            ],
        };

        await exportToExcel(config);
    };

    const handleExportMultipleSheets = async () => {
        const config = {
            fileName: 'multiple_sheets.xlsx',
            sheets: [
                {
                    sheetName: 'Sales Department',
                    columns: [
                        { header: 'Name', key: 'name', width: 20 },
                        { header: 'Salary', key: 'salary', width: 15 },
                    ],
                    data: sampleData.filter(emp => emp.department === 'Sales'),
                    headerStyle: {
                        backgroundColor: 'FFFF6B6B', // Red
                        textColor: 'FFFFFFFF',
                    },
                },
                {
                    sheetName: 'IT Department',
                    columns: [
                        { header: 'Name', key: 'name', width: 20 },
                        { header: 'Salary', key: 'salary', width: 15 },
                    ],
                    data: sampleData.filter(emp => emp.department === 'IT'),
                    headerStyle: {
                        backgroundColor: 'FF4ECDC4', // Teal
                        textColor: 'FFFFFFFF',
                    },
                },
                {
                    sheetName: 'HR Department',
                    columns: [
                        { header: 'Name', key: 'name', width: 20 },
                        { header: 'Salary', key: 'salary', width: 15 },
                    ],
                    data: sampleData.filter(emp => emp.department === 'HR'),
                    headerStyle: {
                        backgroundColor: 'FF95E1D3', // Mint
                        textColor: 'FF000000',
                    },
                },
            ],
        };

        await exportToExcel(config);
    };

    return (
        <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
            <h1>ExcelJS Export Examples</h1>
            <p>Click the buttons below to export Excel files with different features:</p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '15px', maxWidth: '500px' }}>
                <button onClick={handleExportBasic} style={{ ...buttonBaseStyle, backgroundColor: '#4472C4' }}>
                    Export Basic Excel
                </button>

                <button onClick={handleExportWithAllFeatures} style={{ ...buttonBaseStyle, backgroundColor: '#2E75B6' }}>
                    Export with All Features
                </button>

                <button onClick={handleExportMultipleSheets} style={{ ...buttonBaseStyle, backgroundColor: '#00B894' }}>
                    Export Multiple Sheets
                </button>
            </div>

            <div style={{ marginTop: '30px', backgroundColor: '#f5f5f5', padding: '20px', borderRadius: '5px' }}>
                <h2>Features Demonstrated:</h2>
                <ul>
                    <li>✅ <strong>Comments on Headers</strong> - Hover over ID, Department, and Salary header cells</li>
                    <li>✅ <strong>Dropdowns</strong> - Department and Status columns have dropdown lists</li>
                    <li>✅ <strong>Header Styling</strong> - Custom background colors, text colors, and borders</li>
                    <li>✅ <strong>Data Styling</strong> - Apply borders and styles to all data rows</li>
                    <li>✅ <strong>Column Colors</strong> - Status and Salary columns have styling</li>
                    <li>✅ <strong>Row Styles</strong> - Row 3 has green background, rows 5-6 have red background</li>
                    <li>✅ <strong>Cell Styles</strong> - Specific cells have custom styling</li>
                    <li>✅ <strong>Sheet Protection</strong> - Sheet is protected with password</li>
                    <li>✅ <strong>Locked Columns</strong> - ID and Name columns are locked</li>
                    <li>✅ <strong>Freeze Panes</strong> - First row is frozen</li>
                    <li>✅ <strong>Multiple Sheets</strong> - Create workbooks with multiple worksheets</li>
                    <li>✅ <strong>Custom Data</strong> - Add any data structure you want</li>
                </ul>
            </div>

            <div style={{ marginTop: '20px', backgroundColor: '#e3f2fd', padding: '20px', borderRadius: '5px' }}>
                <h3>How to Use in Your Components:</h3>
                <pre style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '5px', overflow: 'auto' }}>
                    {`import exportToExcel from './hooks/useExcelExport';

const MyComponent = () => {
  const handleExport = async () => {
    const config = {
      fileName: 'my_data.xlsx',
      sheets: [
        {
          sheetName: 'Sheet1',
          columns: [
            { header: 'Column1', key: 'col1', width: 15 },
            { header: 'Column2', key: 'col2', width: 20 },
          ],
          data: [
            { col1: 'Value1', col2: 'Value2' },
            { col1: 'Value3', col2: 'Value4' },
          ],
          customData: {
            data: [
              { col1: 'CustomValue1', col2: 'CustomValue2' },
            ],
            style: {
              font: { bold: true, color: { argb: 'FFFF0000' } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEAA7' } },
              height: 25,
            },
          },
          headerStyle: {
            backgroundColor: 'FF4472C4',
            textColor: 'FFFFFFFF',
          },
          dataStyle: {
            border: {
              top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
              left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
              bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
              right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
            },
          },
          freeze: { rows: 1, columns: 0 },
          protectSheet: true, // or { password: 'pass', selectLockedCells: true }
          dropdowns: [
            {
              key: 'col1',
              options: ['Option1', 'Option2'],
            },
          ],
          columnColors: {
            col1: { backgroundColor: 'FFFFEAA7', textColor: 'FF000000' },
          },
          rowStyles: [
            { row: 2, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } } } },
            { startRow: 3, endRow: 5, style: { font: { bold: true } } },
          ],
          cellStyles: [
            { cell: 'A2', style: { font: { color: { argb: 'FFFF0000' } } } },
            { row: 3, column: 'col1', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } } },
          ],
        },
      ],
    };
    await exportToExcel(config);
  };

  return <button onClick={handleExport}>Export</button>;
};`}
                </pre>
            </div>
        </div>
    );
};

export default ExcelExportExample;
