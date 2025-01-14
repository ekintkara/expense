const express = require('express');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// Excel dosyasını okuma
const filePath = path.join(__dirname, 'masraf.xlsx');
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' }); // Boş hücreler için boş string kullan

console.log('Excel Column Names:', Object.keys(sheetData[0]));
console.log('First Row of Data:', sheetData[0]);

// Function to convert Excel data to HTML
function generateHTMLFromExcel(data) {
    let html = `<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Expense Visualization</title>
        <style>
            body {
                padding: 20px;
                font-family: Arial, sans-serif;
                background-color: #f4f4f9;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                background-color: #fff;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            th, td {
                border: 1px solid #ddd;
                padding: 10px;
                text-align: center;
            }
            th {
                background-color: #4CAF50;
                color: white;
            }
            tr:nth-child(even) {
                background-color: #f2f2f2;
            }
            @media screen and (max-width: 600px) { table { display: block; overflow-x: auto; white-space: nowrap; } }
        </style>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    </head>
    <body>
        <h1>Expense Data</h1>`;

    let table1 = data.slice(0, data.findIndex(row => row[Object.keys(data[0])[0]] === 'KARYA BU AY 9 KERE SABAH İŞE GİTTİ'));
    let table2 = data.slice(data.findIndex(row => row[Object.keys(data[0])[0]] === 'KARYA BU AY 9 KERE SABAH İŞE GİTTİ'));

    html += `<table><thead><tr>
        <th>Masraflar</th>
        <th>${Object.keys(table1[0])[1]}</th>
        <th>${Object.keys(table1[0])[2]}</th>
        <th>${Object.keys(table1[0])[3]}</th>
        <th>${Object.keys(table1[0])[4]}</th>
        <th>${Object.keys(table1[0])[5]}</th>
    </tr></thead><tbody>`;

    table1.forEach(row => {
        if (row[Object.keys(table1[0])[0]]) {
            html += `<tr>
                <td>${row[Object.keys(table1[0])[0]]}</td>
                <td>${parseFloat(row[Object.keys(table1[0])[1]]).toFixed(2)}</td>
                <td>${parseFloat(row[Object.keys(table1[0])[2]]).toFixed(2)}</td>
                <td>${parseFloat(row[Object.keys(table1[0])[3]]).toFixed(2)}</td>
                <td>${parseFloat(row[Object.keys(table1[0])[4]]).toFixed(2)}</td>
                <td>${parseFloat(row[Object.keys(table1[0])[5]]).toFixed(2)}</td>
            </tr>`;
        }
    });

    html += `</tbody></table>`;

    html += `<table><thead><tr>
        <th>Toplam Masraf</th>
        <th>Değer</th>
    </tr></thead><tbody>`;

    table2.forEach(row => {
        if (row[Object.keys(table2[0])[0]]) {
            html += `<tr>
                <td>${row[Object.keys(table2[0])[0]]}</td>
                <td>${parseFloat(row[Object.keys(table2[0])[1]]).toFixed(2)}</td>
            </tr>`;
        }
    });

    html += `</tbody></table>
        <div class="container">
            <canvas id="expenseChart"></canvas>
        </div>
        <script>
            const ctx = document.getElementById('expenseChart').getContext('2d');
            const expenseChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: data.map(row => row[Object.keys(data[0])[0]]),
                    datasets: [{
                        label: '${Object.keys(data[0])[4]}',
                        data: data.map(row => parseFloat(row[Object.keys(data[0])[4]])),
                        backgroundColor: ['rgba(75, 192, 192, 0.2)', 'rgba(153, 102, 255, 0.2)', 'rgba(255, 159, 64, 0.2)'],
                        borderColor: ['rgba(75, 192, 192, 1)', 'rgba(153, 102, 255, 1)', 'rgba(255, 159, 64, 1)'],
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        </script>
    </body>
    </html>`;
    return html;
}

// Update the main route to use the generated HTML
app.get('/', (req, res) => {
    const html = generateHTMLFromExcel(sheetData);
    res.send(html);
});

app.listen(port, () => {
    console.log(`Uygulama http://localhost:${port} adresinde çalışıyor.`);
});
