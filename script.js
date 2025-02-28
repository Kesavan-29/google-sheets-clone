document.addEventListener("DOMContentLoaded", () => {
    const tableBody = document.querySelector("#spreadsheet tbody");
    const formulaBar = document.getElementById("formulaBar");
    const exportBtn = document.getElementById("exportCSV");
    const importInput = document.getElementById("importCSV");
    const columns = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
    const rows = 100;
    let cellData = JSON.parse(localStorage.getItem("spreadsheetData")) || {};

    // Create spreadsheet header
    const headerRow = document.querySelector("#spreadsheet thead tr");
    columns.forEach(col => {
        const th = document.createElement("th");
        th.textContent = col;
        headerRow.appendChild(th);
    });

    // Create spreadsheet rows
    for (let i = 1; i <= rows; i++) {
        const row = document.createElement("tr");
        const rowHeader = document.createElement("th");
        rowHeader.textContent = i;
        row.appendChild(rowHeader);

        columns.forEach(col => {
            const cell = document.createElement("td");
            const cellId = `${col}${i}`;
            cell.setAttribute("contenteditable", "true");
            cell.setAttribute("data-cell", cellId);
            cell.textContent = cellData[cellId] || "";

            // Handle input changes
            cell.addEventListener("input", (event) => {
                let value = event.target.textContent.trim();
                if (value.startsWith("=")) {
                    processFormula(cellId, value.substring(1));
                } else {
                    cellData[cellId] = isNaN(value) ? value : Number(value);
                    saveToLocalStorage();
                }
            });

            row.appendChild(cell);
        });
        tableBody.appendChild(row);
    }

    function processFormula(cellId, formula) {
        let match = formula.match(/(SUM|AVERAGE|MAX|MIN|COUNT|UPPER|LOWER)([^)]+)/i);
        if (match) {
            let operation = match[1].toUpperCase();
            let params = match[2].split(",");
            let values = getValuesFromRange(params[0]);
            let result;

            switch (operation) {
                case "SUM": result = values.reduce((a, b) => a + b, 0); break;
                case "AVERAGE": result = values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0; break;
                case "MAX": result = Math.max(...values); break;
                case "MIN": result = Math.min(...values); break;
                case "COUNT": result = values.length; break;
                case "UPPER": result = getCellValue(params[0]).toUpperCase(); break;
                case "LOWER": result = getCellValue(params[0]).toLowerCase(); break;
            }

            cellData[cellId] = result;
            updateCell(cellId, result);
            saveToLocalStorage();
        }
    }

    function getValuesFromRange(range) {
        let [start, end] = range.split(":");
        let values = [];

        if (start && end) {
            let startCol = start.charAt(0);
            let startRow = parseInt(start.substring(1));
            let endRow = parseInt(end.substring(1));

            for (let i = startRow; i <= endRow; i++) {
                let key = `${startCol}${i}`;
                if (cellData[key] !== undefined) {
                    values.push(Number(cellData[key]) || 0);
                }
            }
        }
        return values;
    }

    function getCellValue(cellId) {
        return cellData[cellId] !== undefined ? String(cellData[cellId]) : "";
    }

    function updateCell(cellId, value) {
        let cell = document.querySelector(`[data-cell='${cellId}']`);
        if (cell) cell.textContent = value;
    }

    function saveToLocalStorage() {
        localStorage.setItem("spreadsheetData", JSON.stringify(cellData));
    }

    // Export CSV
    exportBtn.addEventListener("click", () => {
        let csvContent = "data:text/csv;charset=utf-8,";
        csvContent += " ," + columns.join(",") + "\n";
        
        for (let i = 1; i <= rows; i++) {
            let row = [i];
            columns.forEach(col => {
                let cellId = `${col}${i}`;
                row.push(cellData[cellId] || "");
            });
            csvContent += row.join(",") + "\n";
        }

        let encodedUri = encodeURI(csvContent);
        let link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "spreadsheet.csv");
        document.body.appendChild(link);
        link.click();
    });

    // Import CSV
    importInput.addEventListener("change", (event) => {
        let file = event.target.files[0];
        if (!file) return;

        let reader = new FileReader();
        reader.onload = function (e) {
            let rows = e.target.result.split("\n").slice(1);
            rows.forEach((row, i) => {
                let values = row.split(",");
                values.forEach((val, j) => {
                    if (j > 0 && j <= columns.length) {
                        let cellId = `${columns[j - 1]}${i + 1}`;
                        cellData[cellId] = val.trim();
                        updateCell(cellId, val.trim());
                    }
                });
            });
            saveToLocalStorage();
        };
        reader.readAsText(file);
    });
});
