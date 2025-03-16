document.getElementById('analyze-btn').addEventListener('click', function() {
    const fileInput = document.getElementById('excel-file');
    const errorMessage = document.getElementById('error-message');
    errorMessage.style.display = 'none';

    if (fileInput.files.length === 0) {
        errorMessage.textContent = 'Por favor, selecione um arquivo Excel.';
        errorMessage.style.display = 'block';
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Verificar se as abas necessárias existem
        const requiredSheets = ['Vendas', 'Despesas', 'Resultado'];
        const missingSheets = requiredSheets.filter(sheet => !workbook.Sheets[sheet]);

        if (missingSheets.length > 0) {
            errorMessage.textContent = `O arquivo não contém as abas necessárias: ${missingSheets.join(', ')}`;
            errorMessage.style.display = 'block';
            return;
        }

        // Processar a aba "Vendas"
        const vendasSheet = workbook.Sheets['Vendas'];
        const vendasData = XLSX.utils.sheet_to_json(vendasSheet, { header: 1 });
        const totalVendas = displayVendas(vendasData);

        // Processar a aba "Despesas"
        const despesasSheet = workbook.Sheets['Despesas'];
        const despesasData = XLSX.utils.sheet_to_json(despesasSheet, { header: 1 });
        const totalDespesas = displayDespesas(despesasData);

        // Calcular e exibir o resultado final
        const resultadoFinal = totalVendas - totalDespesas;
        document.getElementById('resultado-final').textContent = resultadoFinal.toFixed(2);

        // Gerar gráficos
        generateCharts(vendasData, despesasData);
    };

    reader.readAsArrayBuffer(file);
});

function displayVendas(data) {
    const tbody = document.querySelector('#vendas-table tbody');
    tbody.innerHTML = '';

    let totalVendas = 0;

    // Ignorar a primeira linha (cabeçalho)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tr = document.createElement('tr');
        row.forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);

        // Calcular o total de vendas
        if (row[3]) {
            totalVendas += parseFloat(row[3]);
        }
    }

    document.getElementById('total-vendas').textContent = totalVendas.toFixed(2);
    return totalVendas;
}

function displayDespesas(data) {
    const tbody = document.querySelector('#despesas-table tbody');
    tbody.innerHTML = '';

    let totalDespesas = 0;

    // Ignorar a primeira linha (cabeçalho)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tr = document.createElement('tr');
        row.forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);

        // Calcular o total de despesas
        if (row[1]) {
            totalDespesas += parseFloat(row[1]);
        }
    }

    document.getElementById('total-despesas').textContent = totalDespesas.toFixed(2);
    return totalDespesas;
}

function generateCharts(vendasData, despesasData) {
    const vendasLabels = vendasData.slice(1).map(row => row[0]);
    const vendasValues = vendasData.slice(1).map(row => parseFloat(row[3]));

    const despesasLabels = despesasData.slice(1).map(row => row[0]);
    const despesasValues = despesasData.slice(1).map(row => parseFloat(row[1]));

    // Gráfico de Vendas (Barras)
    new Chart(document.getElementById('vendas-chart'), {
        type: 'bar',
        data: {
            labels: vendasLabels,
            datasets: [{
                label: 'Vendas (R$)',
                data: vendasValues,
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: '#555' // Linhas do grid
                    }
                },
                x: {
                    grid: {
                        color: '#555' // Linhas do grid
                    }
                }
            },
            plugins: {
                legend: {
                    labels: {
                        color: '#fff' // Cor do texto da legenda
                    }
                }
            }
        }
    });

    // Gráfico de Despesas (Pizza)
    new Chart(document.getElementById('despesas-chart'), {
        type: 'pie',
        data: {
            labels: despesasLabels,
            datasets: [{
                label: 'Despesas (R$)',
                data: despesasValues,
                backgroundColor: [
                    'rgba(255, 99, 132, 0.2)',
                    'rgba(54, 162, 235, 0.2)',
                    'rgba(255, 206, 86, 0.2)',
                    'rgba(75, 192, 192, 0.2)',
                    'rgba(153, 102, 255, 0.2)',
                    'rgba(255, 159, 64, 0.2)'
                ],
                borderColor: [
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            plugins: {
                legend: {
                    labels: {
                        color: '#fff' // Cor do texto da legenda
                    }
                }
            }
        }
    });
}
