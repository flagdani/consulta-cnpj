function formatCurrency(value) {
    if (value === null || value === undefined || isNaN(value)) return '-';
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
}

function formatDateBR(dateString) {
    if (!dateString) return '-';
    const date = new Date(dateString);
    if (isNaN(date)) return '-';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function isDataExclusao2023(dateString) {
    if (!dateString) return false;
    const date = new Date(dateString);
    if (isNaN(date)) return false;
    return date.getFullYear() === 2023;
}

// Campos obrigatórios logo após o CNPJ na ordem solicitada
const requiredOrder = [
    "cnpj",
    "opcao_pelo_simples",
    "data_opcao_pelo_simples",
    "data_exclusao_do_simples",
    "codigo_natureza_juridica",
    "data_inicio_atividade",
    "descricao_situacao_cadastral"
];

// Função para obter a ordem final de campos de um objeto de dados,
// garantindo o requiredOrder no início e os demais campos posteriormente em ordem alfabética
function getFieldOrder(data) {
    const allKeys = Object.keys(data);

    // Filtra as chaves obrigatórias
    const mandatory = requiredOrder.filter(key => allKeys.includes(key));

    // Remove obrigatórios e cnpj do conjunto total
    const remaining = allKeys.filter(key => !mandatory.includes(key));

    // Ordena restantes alfabeticamente
    remaining.sort((a,b) => a.localeCompare(b));

    // Combina obrigatórios + restantes
    const finalOrder = [...mandatory, ...remaining];

    return finalOrder;
}

function displaySingleResult(data) {
    const container = document.getElementById('single-result');
    container.innerHTML = '';

    const tableResponsive = document.createElement('div');
    tableResponsive.className = 'table-responsive';

    const table = document.createElement('table');
    table.className = 'table table-striped table-bordered align-middle';

    const tbody = document.createElement('tbody');

    const fields = getFieldOrder(data);

    fields.forEach(key => {
        let value = data[key];

        // Formatações específicas
        if (key === 'capital_social') {
            value = formatCurrency(value);
        } else if (key === 'opcao_pelo_simples') {
            value = value ? 'Sim' : 'Não';
        } else if (key.includes('data')) {
            value = formatDateBR(value);
        } else if (key === 'cnaes_secundarios') {
            if (Array.isArray(value) && value.length > 0) {
                value = value.map(cnae => `Código: ${cnae.codigo || '-'}, Descrição: ${cnae.descricao || '-'}`).join('<br>');
            } else {
                value = '-';
            }
        } else if (key === 'qsa') {
            if (Array.isArray(value) && value.length > 0) {
                value = `<button class="btn btn-info btn-sm view-qsa" data-qsa='${JSON.stringify(value)}'>
                            <i class="fas fa-eye"></i> Ver Detalhes
                         </button>`;
            } else {
                value = '-';
            }
        }

        if (value === null || value === undefined || value === '') {
            value = '-';
        }

        const tr = document.createElement('tr');

        if (key === 'data_exclusao_do_simples' && isDataExclusao2023(data['data_exclusao_do_simples'])) {
            tr.classList.add('highlight-red');
        }

        const th = document.createElement('th');
        th.scope = 'row';
        th.textContent = key.replace(/_/g, ' ').toUpperCase();

        const td = document.createElement('td');
        td.innerHTML = value;

        tr.appendChild(th);
        tr.appendChild(td);
        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableResponsive.appendChild(table);
    container.appendChild(tableResponsive);

    document.getElementById('export-single-excel').style.display = 'inline-block';
    document.getElementById('export-single-excel').dataset.result = JSON.stringify(data);
}

function displayBatchResult(results) {
    const container = document.getElementById('batch-result');
    container.innerHTML = '';

    if (results.length === 0) {
        container.innerHTML = '<p>Nenhum dado encontrado.</p>';
        return;
    }

    // Coletar todos os campos de todos os resultados para saber a ordem completa
    const allKeys = new Set();
    results.forEach(r => {
        if (r.data) {
            Object.keys(r.data).forEach(k => allKeys.add(k));
        }
    });
    allKeys.add('Erro');

    // Converter para array
    let keysArray = Array.from(allKeys);
    // Remover "Erro" temporariamente para ordenar
    const hasErro = keysArray.includes('Erro');
    if (hasErro) {
        keysArray = keysArray.filter(k => k !== 'Erro');
    }

    // Aplicar a ordem obrigatória nos campos (exceto "Erro")
    // Como precisamos da ordem do primeiro resultado para ter referência, vamos pegar o primeiro data válido
    let referenceData = results.find(r => r.data !== null && r.data !== undefined);
    if (!referenceData) {
        // Se não há nenhum data válido, apenas ordena alfabeticamente + erro
        keysArray.sort((a,b) => a.localeCompare(b));
    } else {
        const finalOrder = getFieldOrder(referenceData.data);
        // Adicionar chaves extras que não estejam no referenceData
        const extras = keysArray.filter(k => !finalOrder.includes(k)).sort((a,b) => a.localeCompare(b));
        keysArray = [...finalOrder, ...extras];
    }

    // Re-adicionar "Erro" no final
    if (hasErro) {
        keysArray.push('Erro');
    }

    const tableResponsive = document.createElement('div');
    tableResponsive.className = 'table-responsive';

    const table = document.createElement('table');
    table.className = 'table table-striped table-bordered align-middle';

    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    const trHead = document.createElement('tr');
    keysArray.forEach(header => {
        const th = document.createElement('th');
        th.scope = 'col';
        th.textContent = header.replace(/_/g, ' ').toUpperCase();
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    results.forEach(result => {
        const tr = document.createElement('tr');
        keysArray.forEach(header => {
            const td = document.createElement('td');
            if (header === 'Erro') {
                td.textContent = result.error || '-';
            } else if (result.error) {
                td.textContent = '-';
            } else {
                let value = result.data[header];
                if (header === 'capital_social') {
                    value = formatCurrency(value);
                } else if (header === 'opcao_pelo_simples') {
                    value = value ? 'Sim' : 'Não';
                } else if (header.includes('data')) {
                    value = formatDateBR(value);
                } else if (header === 'cnaes_secundarios') {
                    if (Array.isArray(value) && value.length > 0) {
                        value = value.map(cnae => `Código: ${cnae.codigo || '-'}, Descrição: ${cnae.descricao || '-'}`).join('<br>');
                    } else {
                        value = '-';
                    }
                } else if (header === 'qsa') {
                    if (Array.isArray(value) && value.length > 0) {
                        value = `<button class="btn btn-info btn-sm view-qsa" data-qsa='${JSON.stringify(value)}'>
                                    <i class="fas fa-eye"></i> Ver Detalhes
                                 </button>`;
                    } else {
                        value = '-';
                    }
                }

                if (value === null || value === undefined || value === '') {
                    value = '-';
                }

                td.innerHTML = value;

                if (header === 'data_exclusao_do_simples' && isDataExclusao2023(result.data['data_exclusao_do_simples'])) {
                    td.classList.add('highlight-red');
                }
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    tableResponsive.appendChild(table);
    container.appendChild(tableResponsive);
}

function exportToExcel(results) {
    if (results.length === 0) {
        alert('Nenhum dado para exportar.');
        return;
    }

    // Obter campos
    const allKeys = new Set();
    results.forEach(result => {
        if (result.data) {
            Object.keys(result.data).forEach(key => allKeys.add(key));
        }
    });
    allKeys.add('Erro');

    let keysArray = Array.from(allKeys);
    const hasErro = keysArray.includes('Erro');
    if (hasErro) {
        keysArray = keysArray.filter(k => k !== 'Erro');
    }

    let referenceData = results.find(r => r.data !== null && r.data !== undefined);
    if (!referenceData) {
        // Sem dados válidos, ordenar tudo alfabeticamente
        keysArray.sort((a,b) => a.localeCompare(b));
    } else {
        const finalOrder = getFieldOrder(referenceData.data);
        const extras = keysArray.filter(k => !finalOrder.includes(k)).sort((a,b) => a.localeCompare(b));
        keysArray = [...finalOrder, ...extras];
    }

    if (hasErro) keysArray.push('Erro');

    const wb = XLSX.utils.book_new();
    const ws_data = [];

    ws_data.push(keysArray.map(header => header.replace(/_/g, ' ').toUpperCase()));

    const dataRows = [];
    results.forEach(result => {
        const row = [];
        keysArray.forEach(header => {
            if (header === 'Erro') {
                row.push(result.error || '-');
            } else if (result.error) {
                row.push('-');
            } else {
                let value = result.data[header];
                if (header === 'capital_social') {
                    value = formatCurrency(value);
                } else if (header === 'opcao_pelo_simples') {
                    value = value ? 'Sim' : 'Não';
                } else if (header === 'cnaes_secundarios') {
                    if (Array.isArray(value) && value.length > 0) {
                        value = value.map(cnae => `Código: ${cnae.codigo || '-'}, Descrição: ${cnae.descricao || '-'}`).join('; ');
                    } else {
                        value = '-';
                    }
                } else if (header === 'qsa') {
                    if (Array.isArray(value) && value.length > 0) {
                        value = value.map(socio =>
`Nome: ${socio.nome_socio || '-'}
CPF/CNPJ: ${socio.cnpj_cpf_do_socio || '-'}
Qualificação: ${socio.qualificacao_socio || '-'}
Data de Entrada: ${formatDateBR(socio.data_entrada_sociedade)}`).join('\n\n');
                    } else {
                        value = '-';
                    }
                } else if (header.includes('data')) {
                    value = formatDateBR(value);
                }

                if (value === null || value === undefined || value === '') {
                    value = '-';
                }

                row.push(value);
            }
        });
        dataRows.push({ row, originalData: result.data });
        ws_data.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");

    const dataExclusaoIndex = keysArray.indexOf('data_exclusao_do_simples');
    if (dataExclusaoIndex !== -1) {
        for (let i = 0; i < dataRows.length; i++) {
            const originalData = dataRows[i].originalData;
            if (originalData && isDataExclusao2023(originalData['data_exclusao_do_simples'])) {
                const cell_ref = XLSX.utils.encode_cell({ c: dataExclusaoIndex, r: i + 1 });
                if (!ws[cell_ref]) continue;
                ws[cell_ref].s = {
                    fill: { fgColor: { rgb: "FFE5E5" } }
                };
            }
        }
    }

    const headerRange = XLSX.utils.decode_range(ws['!ref']);
    for (let C = headerRange.s.c; C <= headerRange.e.c; ++C) {
        const cell_address = { c: C, r: 0 };
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        if (!ws[cell_ref]) continue;
        ws[cell_ref].s = {
            font: { bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "4F81BD" } }
        };
    }

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'resultados_consulta_cnpj.xlsx');
}

function exportSingleToExcel(data) {
    const wb = XLSX.utils.book_new();

    const fields = getFieldOrder(data);

    const mainData = [];
    mainData.push(["Campo", "Valor"]);

    fields.forEach(key => {
        let value = data[key];
        if (key === 'capital_social') {
            value = formatCurrency(value);
        } else if (key === 'opcao_pelo_simples') {
            value = value ? 'Sim' : 'Não';
        } else if (key.includes('data')) {
            value = formatDateBR(value);
        }

        if (key === 'cnaes_secundarios') {
            if (Array.isArray(value) && value.length > 0) {
                value = value.map(cnae => `Código: ${cnae.codigo || '-'}, Descrição: ${cnae.descricao || '-'}`).join('; ');
            } else {
                value = '-';
            }
        }

        if (value === null || value === undefined || value === '') {
            value = '-';
        }

        mainData.push([key.replace(/_/g, ' ').toUpperCase(), value]);
    });

    const ws_main = XLSX.utils.aoa_to_sheet(mainData);
    XLSX.utils.book_append_sheet(wb, ws_main, "Dados Principais");

    const dataExclusaoRow = fields.indexOf('data_exclusao_do_simples') + 1; 
    if (dataExclusaoRow > 0 && isDataExclusao2023(data['data_exclusao_do_simples'])) {
        const cell_ref = XLSX.utils.encode_cell({ c: 1, r: dataExclusaoRow });
        if (ws_main[cell_ref]) {
            ws_main[cell_ref].s = {
                fill: { fgColor: { rgb: "FFE5E5" } }
            };
        }
    }

    if (Array.isArray(data.qsa) && data.qsa.length > 0) {
        const qsaData = [];
        const qsaHeaders = Object.keys(data.qsa[0]);
        qsaData.push(qsaHeaders.map(header => header.replace(/_/g, ' ').toUpperCase()));

        data.qsa.forEach(entry => {
            const row = [];
            qsaHeaders.forEach(header => {
                let value = entry[header];
                if (header.includes('data')) {
                    value = formatDateBR(value);
                }
                if (value === null || value === undefined || value === '') {
                    value = '-';
                }
                row.push(value);
            });
            qsaData.push(row);
        });

        const ws_qsa = XLSX.utils.aoa_to_sheet(qsaData);
        XLSX.utils.book_append_sheet(wb, ws_qsa, "QSA");
    }

    if (Array.isArray(data.cnaes_secundarios) && data.cnaes_secundarios.length > 0) {
        const atividadesData = [];
        atividadesData.push(["Código", "Descrição"]);
        data.cnaes_secundarios.forEach(cnae => {
            atividadesData.push([cnae.codigo || '-', cnae.descricao || '-']);
        });

        const ws_atividades = XLSX.utils.aoa_to_sheet(atividadesData);
        XLSX.utils.book_append_sheet(wb, ws_atividades, "Atividades Secundárias");
    }

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), `consulta_cnpj_${data.cnpj}.xlsx`);
}

function showQsaModal(qsaData) {
    const modal = new bootstrap.Modal(document.getElementById('qsaModal'));
    const content = document.getElementById('qsa-content');
    content.innerHTML = '';

    if (Array.isArray(qsaData) && qsaData.length > 0) {
        const table = document.createElement('table');
        table.className = 'table table-striped table-bordered align-middle';

        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        const headers = Object.keys(qsaData[0]);
        const trHead = document.createElement('tr');
        headers.forEach(header => {
            const th = document.createElement('th');
            th.scope = 'col';
            th.textContent = header.replace(/_/g, ' ').toUpperCase();
            trHead.appendChild(th);
        });
        thead.appendChild(trHead);

        qsaData.forEach(entry => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                let value = entry[header];
                if (header.includes('data')) {
                    value = formatDateBR(value);
                }
                if (value === null || value === undefined || value === '') {
                    value = '-';
                }
                td.textContent = value;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        table.appendChild(thead);
        table.appendChild(tbody);
        content.appendChild(table);
    } else {
        content.textContent = 'Nenhum dado de QSA disponível.';
    }

    modal.show();
}

function updateProgress(current, total) {
    const progressBar = document.getElementById('progress-bar');
    const progressContainer = document.getElementById('batch-progress');
    const progressStatus = document.getElementById('progress-status');

    const percent = Math.round((current / total) * 100);
    progressBar.style.width = percent + '%';
    progressBar.setAttribute('aria-valuenow', percent);
    progressBar.textContent = `${percent}%`;

    progressStatus.textContent = `Processado ${current} de ${total} CNPJs.`;
}

async function fetchSingleCNPJ() {
    const cnpjInput = document.getElementById('single-cnpj');
    const cnpj = cnpjInput.value.trim().replace(/\D/g, '');
    const resultContainer = document.getElementById('single-result');
    const errorContainer = document.getElementById('single-error');
    const loader = document.getElementById('single-loader');
    const exportBtn = document.getElementById('export-single-excel');
    resultContainer.innerHTML = '';
    errorContainer.textContent = '';
    exportBtn.style.display = 'none';

    if (!/^\d{14}$/.test(cnpj)) {
        errorContainer.textContent = 'CNPJ inválido. Deve conter 14 dígitos numéricos.';
        return;
    }

    loader.style.display = 'inline-block';
    try {
        const response = await fetch(`https://brasilapi.com.br/api/cnpj/v1/${cnpj}`);
        if (!response.ok) {
            throw new Error(`Erro na requisição: ${response.statusText}`);
        }
        const data = await response.json();
        displaySingleResult(data);
    } catch (error) {
        errorContainer.textContent = error.message;
    } finally {
        loader.style.display = 'none';
    }
}

async function fetchBatchCNPJ() {
    const fileInput = document.getElementById('excel-file');
    const file = fileInput.files[0];
    const resultContainer = document.getElementById('batch-result');
    const errorContainer = document.getElementById('batch-error');
    const loader = document.getElementById('batch-loader');
    const exportBtn = document.getElementById('export-excel');
    const progressBar = document.getElementById('batch-progress');
    const progressStatus = document.getElementById('progress-status');
    resultContainer.innerHTML = '';
    errorContainer.textContent = '';
    exportBtn.style.display = 'none';
    progressBar.style.display = 'none';
    progressStatus.textContent = '';

    if (!file) {
        errorContainer.textContent = 'Por favor, selecione um arquivo Excel.';
        return;
    }

    const reader = new FileReader();
    reader.onload = async function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

        if (json.length === 0) {
            errorContainer.textContent = 'Arquivo Excel vazio.';
            return;
        }

        const headers = Object.keys(json[0]).map(h => h.trim().toLowerCase().replace(/\s/g, ''));
        const cnpjIndex = headers.findIndex(header => header === 'cnpj');

        if (cnpjIndex === -1) {
            errorContainer.textContent = 'Coluna "CNPJ" não encontrada no arquivo Excel.';
            return;
        }

        const cnpjs = json.map(row => {
            const key = Object.keys(row)[cnpjIndex];
            return String(row[key]).trim().replace(/\D/g, '');
        }).filter(cnpj => cnpj !== '');

        if (cnpjs.length === 0) {
            errorContainer.textContent = 'Nenhum CNPJ válido encontrado no arquivo Excel.';
            return;
        }

        loader.style.display = 'inline-block';
        progressBar.style.display = 'block';
        updateProgress(0, cnpjs.length);

        const results = [];

        for (let i = 0; i < cnpjs.length; i++) {
            const cnpj = cnpjs[i];
            if (!/^\d{14}$/.test(cnpj)) {
                results.push({ data: null, cnpj, error: 'CNPJ inválido.' });
                updateProgress(i + 1, cnpjs.length);
                continue;
            }
            try {
                const response = await fetch(`https://brasilapi.com.br/api/cnpj/v1/${cnpj}`);
                if (!response.ok) {
                    throw new Error(`Erro: ${response.statusText}`);
                }
                const data = await response.json();
                results.push({ data, cnpj, error: null });
            } catch (error) {
                results.push({ data: null, cnpj, error: error.message });
            }
            updateProgress(i + 1, cnpjs.length);
        }

        displayBatchResult(results);
        loader.style.display = 'none';

        if (results.length > 0) {
            exportBtn.style.display = 'inline-block';
            exportBtn.dataset.results = JSON.stringify(results);
        }

        progressStatus.textContent = 'Processamento concluído.';
    };

    reader.onerror = function() {
        errorContainer.textContent = 'Erro ao ler o arquivo Excel.';
    };

    reader.readAsArrayBuffer(file);
}

document.getElementById('fetch-single').addEventListener('click', fetchSingleCNPJ);
document.getElementById('fetch-batch').addEventListener('click', fetchBatchCNPJ);

document.getElementById('export-excel').addEventListener('click', function() {
    const exportBtn = this;
    const results = JSON.parse(exportBtn.dataset.results);
    exportToExcel(results);
});

document.getElementById('export-single-excel').addEventListener('click', function() {
    const exportBtn = this;
    const data = JSON.parse(exportBtn.dataset.result);
    exportSingleToExcel(data);
});

document.addEventListener('click', function(event) {
    if (event.target && event.target.classList.contains('view-qsa')) {
        const qsaData = JSON.parse(event.target.getAttribute('data-qsa'));
        showQsaModal(qsaData);
    }
});
