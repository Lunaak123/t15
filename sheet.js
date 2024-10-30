let data = []; // Holds the initial Excel data
let filteredData = []; // Holds the filtered data after user operations
let subsheetNames = []; // Holds the names of subsheets

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

        // Load the first sheet data
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Populate subsheet names if any
        subsheetNames = workbook.SheetNames.filter(name => name !== sheetName);
        const subsheetSelect = document.getElementById('subsheet-select');
        subsheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            subsheetSelect.appendChild(option);
        });

        // Convert the sheet to JSON
        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the first sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear previous content

    if (sheetData.length > 0) {
        const table = document.createElement('table');
        const headerRow = document.createElement('tr');
        
        // Create header row
        Object.keys(sheetData[0]).forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });
        table.appendChild(headerRow);

        // Create data rows
        sheetData.forEach(row => {
            const tr = document.createElement('tr');
            Object.values(row).forEach(cell => {
                const td = document.createElement('td');
                td.textContent = cell === null ? 'N/A' : cell; // Handle null values
                tr.appendChild(td);
            });
            table.appendChild(tr);
        });

        sheetContentDiv.appendChild(table);
    } else {
        sheetContentDiv.innerHTML = '<p>No data available.</p>';
    }
}

// Function to apply operations based on user input
function applyOperations() {
    const primaryColumn = document.getElementById('primary-column').value;
    const operationColumns = document.getElementById('operation-columns').value.split(',');
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    filteredData = data.filter(row => {
        const primaryValue = row[primaryColumn] || null;
        const operationsPassed = operationColumns.every(col => {
            const value = row[col] || null;
            return operation === 'null' ? value === null : value !== null;
        });
        return operationsPassed;
    });

    // Display filtered data
    displaySheet(filteredData);
}

// Function to handle subsheet selection
function handleSubsheetSelect() {
    const selectedSubsheet = document.getElementById('subsheet-select').value;
    if (selectedSubsheet) {
        // Load the selected subsheet data
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[selectedSubsheet];
        const subsheetData = XLSX.utils.sheet_to_json(sheet, { defval: null });
        displaySheet(subsheetData);
    } else {
        displaySheet(data); // If no subsheet is selected, show the main data
    }
}

// Function to download filtered data
function downloadFile() {
    const filename = document.getElementById('filename').value || 'download';
    const format = document.getElementById('file-format').value;

    if (format === 'xlsx') {
        const ws = XLSX.utils.json_to_sheet(filteredData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// Event Listeners
document.getElementById('apply-operation').addEventListener('click', applyOperations);
document.getElementById('subsheet-select').addEventListener('change', handleSubsheetSelect);
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex'; // Show modal
});
document.getElementById('confirm-download').addEventListener('click', downloadFile);
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none'; // Hide modal
});

// Load the initial Excel file (replace with the actual URL)
const excelFileUrl = 'path/to/your/excel/file.xlsx'; // Update this to your Excel file URL
loadExcelSheet(excelFileUrl);
