function importExcel() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    if (!file) {
        alert("Please select an Excel file first.");
        return;
    }

    const reader = new FileReader();

    reader.onload = (event) => {
        document.getElementById('loadingMessage').style.display = 'block'; // Show loading message
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON format (array of arrays)
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Display the data
            displayExcelData(jsonData);

            // Generate column options
            generateColumnSelection(jsonData);

            document.getElementById('loadingMessage').style.display = 'none'; // Hide loading message
        } catch (error) {
            console.error("Error reading Excel file:", error);
            alert("There was an error reading the Excel file. Please check the file format.");
        }
    };

    reader.onerror = (error) => {
        console.error("FileReader error:", error);
        alert("There was an error reading the file.");
    };

    reader.readAsArrayBuffer(file);
}

function displayExcelData(data) {
    const excelDataContainer = document.getElementById('excelData');
    let table = '<table border="1" cellpadding="8">';

    data.forEach((row, rowIndex) => {
        table += '<tr>';
        row.forEach((cell) => {
            if (rowIndex === 0) {
                table += `<th>${cell !== undefined ? cell : ''}</th>`;
            } else {
                table += `<td>${cell !== undefined ? cell : 'N/A'}</td>`;
            }
        });
        table += '</tr>';
    });

    table += '</table>';
    excelDataContainer.innerHTML = table;
}

function generateColumnSelection(data) {
    const columnSelection = document.getElementById('columnSelection');
    const xColumnSelect = document.getElementById('xColumn');
    const yColumnSelect = document.getElementById('yColumn');

    const columns = data[0];

    // Clear previous options
    xColumnSelect.innerHTML = '';
    yColumnSelect.innerHTML = '';

    // Populate dropdown with column headers (first row)
    columns.forEach((column, index) => {
        let option = document.createElement("option");
        option.value = index;
        option.textContent = column;
        xColumnSelect.appendChild(option);

        option = document.createElement("option");
        option.value = index;
        option.textContent = column;
        yColumnSelect.appendChild(option);
    });

    columnSelection.style.display = 'block'; // Show the column selection form
}

function drawChart() {
    const xColumn = parseInt(document.getElementById('xColumn').value);
    const yColumn = parseInt(document.getElementById('yColumn').value);

    if (isNaN(xColumn) || isNaN(yColumn)) {
        alert("Please select valid columns for X and Y.");
        return;
    }

    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    const reader = new FileReader();

    reader.onload = (event) => {
        document.getElementById('loadingMessage').style.display = 'block'; // Show loading message
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON format (array of arrays)
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Extract the selected columns data
            const labels = jsonData.slice(1).map(row => row[xColumn]);
            const values = jsonData.slice(1).map(row => parseFloat(row[yColumn])).filter(value => !isNaN(value));

            if (values.length === 0) {
                alert("Selected Y-axis column does not contain valid numeric data.");
                return;
            }

            // Draw the chart with the data
            const chartType = document.getElementById('chartType').value;
            drawChartVisualization(labels, values, chartType);
        } catch (error) {
            console.error("Error reading Excel file:", error);
            alert("There was an error reading the Excel file. Please check the file format.");
        } finally {
            document.getElementById('loadingMessage').style.display = 'none'; // Hide loading message
        }
    };

    reader.onerror = (error) => {
        console.error("FileReader error:", error);
        alert("There was an error reading the file.");
    };

    reader.readAsArrayBuffer(file);
}

function drawChartVisualization(labels, values, chartType) {
    const chartContainer = document.getElementById('myChart').getContext('2d');

    // Destroy any existing chart instance
    if (window.myChartInstance) {
        window.myChartInstance.destroy();
    }

    // Define custom color gradient for line and bar charts
    const gradient = chartContainer.createLinearGradient(0, 0, 0, 400);
    gradient.addColorStop(0, '#000080');  // Navy blue
    gradient.addColorStop(1, '#fcc200');  // Gold

    window.myChartInstance = new Chart(chartContainer, {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: `${chartType.charAt(0).toUpperCase() + chartType.slice(1)} Chart`,
                data: values,
                backgroundColor: chartType === 'pie' ? generateColors(labels.length) : gradient,
                borderColor: '#000080', // Navy blue for border
                borderWidth: 2,
                fill: chartType === 'line' || chartType === 'bar',
                borderRadius: 5,  // Add rounded corners to the bars
            }],
        },
        options: {
            responsive: false,  // Disable dynamic resizing
            maintainAspectRatio: false,  // Allow manual control over the chart size
            plugins: {
                legend: {
                    position: 'top', // Position the legend at the top
                    labels: {
                        font: {
                            size: 16,  // Larger font size for better readability
                            color: '#000080', // Navy blue text for legend
                        },
                    },
                },
                title: {
                    display: true, // Display the chart title
                    text: `Chart for ${labels.length} Items`, // Title can be dynamic based on data
                    font: {
                        size: 18,
                        weight: 'bold',
                        color: '#000080', // Navy blue color for title
                    },
                    padding: 20,
                },
            },
            scales: chartType !== 'pie' ? {
                x: {
                    title: { 
                        display: true, 
                        text: 'X-axis Labels', 
                        font: { size: 14, color: '#000080' } 
                    },
                    ticks: {
                        font: {
                            size: 12,  // Adjust x-axis tick labels size
                            color: '#000080', // Navy blue color for x-axis ticks
                        },
                    },
                },
                y: {
                    title: { 
                        display: true, 
                        text: 'Y-axis Values', 
                        font: { size: 14, color: '#000080' } 
                    },
                    ticks: {
                        font: {
                            size: 12,  // Adjust y-axis tick labels size
                            color: '#000080', // Navy blue color for y-axis ticks
                        },
                    },
                },
            } : {},
            layout: {
                padding: {
                    left: 20,
                    right: 20,
                    top: 20,
                    bottom: 20,
                },
            },
            elements: {
                point: {
                    radius: 0, // Remove points from line charts
                },
            },
        }
    });
}

function generateColors(count) {
    const colors = [];
    for (let i = 0; i < count; i++) {
        const r = Math.floor(Math.random() * 255);
        const g = Math.floor(Math.random() * 255);
        const b = Math.floor(Math.random() * 255);
        colors.push(`rgba(${r}, ${g}, ${b}, 0.7)`);
    }
    return colors;
}
