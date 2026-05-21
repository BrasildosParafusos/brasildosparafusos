document.addEventListener('DOMContentLoaded', () => {
    // --- DOM Elements ---
    const fileUpload = document.getElementById('file-upload');
    const welcomeView = document.getElementById('welcome-view');
    const dashboardView = document.getElementById('dashboard-view');
    const loadingOverlay = document.getElementById('loading-overlay');

    // KPI Elements
    const totalItemsEl = document.getElementById('total-items');
    const totalItemsValueEl = document.getElementById('total-items-value');
    const replenishedCountEl = document.getElementById('replenished-count');
    const replenishedValueEl = document.getElementById('replenished-value');
    const nonReplenishedCountEl = document.getElementById('non-replenished-count');
    const nonReplenishedValueEl = document.getElementById('non-replenished-value');
    const replenishRateEl = document.getElementById('replenish-rate');

    // Table & Filters
    const giroTableBody = document.getElementById('giro-table-body');
    const simpleSearch = document.getElementById('simple-search');
    const advancedSearch = document.getElementById('advanced-search');
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const filterBtns = document.querySelectorAll('.filters .filter-btn[data-filter]');

    // Sort Headers
    const sortCodeHeader = document.getElementById('sort-code');
    const sortDescHeader = document.getElementById('sort-desc');
    const sortQtyHeader = document.getElementById('sort-qty');
    const sortValueHeader = document.getElementById('sort-value');
    const sortRecorrenciaHeader = document.getElementById('sort-recorrencia');

    // --- State Management ---
    let rawData = [];
    let processedData = [];
    let filteredData = [];
    let activeFilter = 'all'; // 'all', 'sim', 'nao'
    let sortColumn = 'none'; // 'code', 'desc', 'qty', 'value'
    let sortDirection = 'none'; // 'asc', 'desc'

    // Chart instances
    let doughnutChart = null;
    let supplierChart = null;
    let itemChart = null;

    // --- Utility Helper Functions ---
    const formatBRL = (val) => val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

    function showLoading() { loadingOverlay.style.display = 'flex'; }
    function hideLoading() { loadingOverlay.style.display = 'none'; }

    /**
     * Safely parses values containing Brazilian currency symbols, periods, and commas.
     */
    function parseNumber(val) {
        if (val === null || val === undefined) return 0;
        if (typeof val === 'number') return val;
        let clean = val.toString()
            .replace(/R\$/g, '')
            .replace(/\s/g, '')
            .replace(/\./g, '')
            .replace(/,/g, '.')
            .trim();
        let num = parseFloat(clean);
        return isNaN(num) ? 0 : num;
    }

    /**
     * Safely parses percentage values (e.g. 0.85, "85%", "85,5%")
     */
    function parsePercentage(val) {
        if (val === null || val === undefined) return 0;
        if (typeof val === 'number') {
            if (val > 0 && val <= 1) {
                return val * 100;
            }
            return val;
        }
        let str = val.toString().trim();
        let hasPercent = str.includes('%');
        let clean = str.replace(/%/g, '').replace(/,/g, '.').trim();
        let num = parseFloat(clean);
        if (isNaN(num)) return 0;
        if (hasPercent) return num;
        if (!hasPercent && num > 0 && num <= 1) {
            return num * 100;
        }
        return num;
    }

    /**
     * Generates a beautiful HTML badge for the recurrence percentage
     */
    function getRecorrenciaBadge(val) {
        if (val === undefined || val === null) return `<span class="badge-recur-low">N/D</span>`;
        if (val === 0) return `<span class="badge-recur-low">0%</span>`;
        let colorClass = 'badge-recur-low';
        let icon = '⚪';
        if (val >= 80) {
            colorClass = 'badge-recur-high';
            icon = '🔥';
        } else if (val >= 50) {
            colorClass = 'badge-recur-med';
            icon = '⚡';
        } else {
            colorClass = 'badge-recur-low';
            icon = '📉';
        }
        return `<span class="${colorClass}">${icon} ${val.toFixed(1)}%</span>`;
    }

    /**
     * Case-insensitive robust property finder in excel rows
     */
    function getValueByKeys(row, keys) {
        for (let k of Object.keys(row)) {
            const normalizedK = k.toString().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            for (let targetKey of keys) {
                const normalizedTarget = targetKey.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                if (normalizedK === normalizedTarget || normalizedK.includes(normalizedTarget)) {
                    return row[k];
                }
            }
        }
        return null;
    }

    // --- File Handler & Parsing ---
    if (fileUpload) {
        fileUpload.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            showLoading();

            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { defval: null, blankrows: false });

                    console.log('JSON Giro parsed successfully:', json.length, 'rows.');

                    if (json.length === 0) {
                        alert('A planilha importada parece estar vazia ou mal formatada.');
                        hideLoading();
                        return;
                    }

                    rawData = json;
                    processData(json);
                    
                    // Toggle Views
                    welcomeView.style.display = 'none';
                    dashboardView.style.display = 'block';

                    // Apply filters and draw UI
                    applyFiltersAndSort();
                    initCharts();

                    setTimeout(hideLoading, 500);
                } catch (err) {
                    console.error('Error parsing 6-month spreadsheet:', err);
                    alert('Ocorreu um erro ao processar a planilha. Verifique se as colunas estão corretas.');
                    hideLoading();
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Map Excel columns to our clean dataset
     */
    function processData(rows) {
        processedData = rows.map((row, index) => {
            // Find columns dynamically with fallback arrays
            const codeVal = getValueByKeys(row, ['Itens', 'Código', 'Cod', 'ID', 'Item', 'Produto']);
            const descVal = getValueByKeys(row, ['Descrição', 'Descricao', 'Desc', 'Nome', 'Produto']);
            const unVal = getValueByKeys(row, ['UN', 'Unidade', 'Medida', 'Un']);
            const compradoVal = getValueByKeys(row, ['Comprado', 'Se comprado', 'Comprado 6 meses', 'Giro', 'Movimentado', 'Comprado nos ultimos']);
            const qtyVal = getValueByKeys(row, ['QNT comprada', 'Quantidade comprada', 'Quantidade', 'Qtd', 'Volume', 'QNT']);
            const supplierVal = getValueByKeys(row, ['Fornecedor', 'Fornecedores', 'Fabricante', 'Distribuidor']);
            const valorVal = getValueByKeys(row, ['Valor', 'Valor comprado', 'Preço', 'Total', 'Gasto', 'Valor Total']);
            const recorrenciaVal = getValueByKeys(row, ['Recorrência', 'Recorrencia', 'Recorrente', 'Freqüência', 'Frequencia', 'Recurrency', 'Recurrent', '%']);

            // Parse columns safely
            const code = codeVal !== null ? codeVal.toString().trim() : `W${index + 1}`;
            const desc = descVal !== null ? descVal.toString().trim() : 'Sem descrição cadastrada';
            const un = unVal !== null ? unVal.toString().trim() : 'PC';
            
            // Purchased condition: contains "sim", "yes", true, or "s"
            let comprado = false;
            if (compradoVal !== null) {
                const cStr = compradoVal.toString().trim().toLowerCase();
                comprado = cStr === 'sim' || cStr === 'yes' || cStr === 's' || cStr === 'true' || cStr === '1';
            }

            const qty = parseNumber(qtyVal);
            
            let fornecedor = 'Não informado';
            if (supplierVal !== null) {
                const sStr = supplierVal.toString().trim();
                fornecedor = sStr === '#N/D' || sStr === '' ? 'Sem fornecedor' : sStr;
            }

            const valor = parseNumber(valorVal);
            const recorrencia = parsePercentage(recorrenciaVal);

            return { code, desc, un, comprado, qty, fornecedor, valor, recorrencia };
        });
        
        filteredData = [...processedData];
    }

    // --- Metrics Updating ---
    function updateKPIs() {
        const totalItems = processedData.length;
        const totalValue = processedData.reduce((sum, item) => sum + item.valor, 0);

        const repItems = processedData.filter(i => i.comprado);
        const repCount = repItems.length;
        const repValue = repItems.reduce((sum, item) => sum + item.valor, 0);

        const nonRepItems = processedData.filter(i => !i.comprado);
        const nonRepCount = nonRepItems.length;
        const nonRepValue = nonRepItems.reduce((sum, item) => sum + item.valor, 0);

        const rate = totalItems > 0 ? (repCount / totalItems) * 100 : 0;
        const avgRecurrence = repCount > 0 
            ? repItems.reduce((sum, item) => sum + (item.recorrencia || 0), 0) / repCount
            : 0;

        // Render to DOM
        totalItemsEl.textContent = totalItems.toLocaleString('pt-BR');
        totalItemsValueEl.textContent = `Giro total: ${formatBRL(totalValue)}`;

        replenishedCountEl.textContent = repCount.toLocaleString('pt-BR');
        replenishedValueEl.textContent = `Total gasto: ${formatBRL(repValue)}`;

        nonReplenishedCountEl.textContent = nonRepCount.toLocaleString('pt-BR');
        nonReplenishedValueEl.textContent = `Gasto retido: ${formatBRL(nonRepValue)}`;

        replenishRateEl.textContent = `${rate.toFixed(1)}%`;

        const avgRecurrenceEl = document.getElementById('avg-recurrence');
        if (avgRecurrenceEl) {
            avgRecurrenceEl.textContent = `${avgRecurrence.toFixed(1)}%`;
        }
    }

    // --- Search & Filters implementation ---
    function applyFiltersAndSort() {
        const term = simpleSearch.value.toLowerCase().trim();
        const smartTerm = advancedSearch.value.toLowerCase().trim();
        const smartKeywords = smartTerm.split(/\s+/).filter(k => k.length > 0);

        // Filter data
        filteredData = processedData.filter(item => {
            // Type Filter (Sim / Não)
            if (activeFilter === 'sim' && !item.comprado) return false;
            if (activeFilter === 'nao' && item.comprado) return false;

            // Simple Search
            if (term.length > 0) {
                const matchCode = item.code.toLowerCase().includes(term);
                const matchDesc = item.desc.toLowerCase().includes(term);
                if (!matchCode && !matchDesc) return false;
            }

            // Advanced Multi-Keyword Search
            if (smartKeywords.length > 0) {
                const matchSmart = smartKeywords.every(kw => 
                    item.code.toLowerCase().includes(kw) || 
                    item.desc.toLowerCase().includes(kw) || 
                    item.fornecedor.toLowerCase().includes(kw)
                );
                if (!matchSmart) return false;
            }

            return true;
        });

        // Apply Sorting
        if (sortColumn !== 'none' && sortDirection !== 'none') {
            filteredData.sort((a, b) => {
                let valA, valB;
                if (sortColumn === 'code') {
                    valA = a.code; valB = b.code;
                    return sortDirection === 'asc' ? valA.localeCompare(valB, undefined, { numeric: true }) : valB.localeCompare(valA, undefined, { numeric: true });
                } else if (sortColumn === 'desc') {
                    valA = a.desc; valB = b.desc;
                    return sortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
                } else if (sortColumn === 'recorrencia') {
                    valA = a.recorrencia || 0; valB = b.recorrencia || 0;
                } else if (sortColumn === 'qty') {
                    valA = a.qty; valB = b.qty;
                } else if (sortColumn === 'value') {
                    valA = a.valor; valB = b.valor;
                }

                if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
                if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
                return 0;
            });
        }

        updateKPIs();
        renderTable();
        updateChartsData();
    }

    // --- Table Rendering ---
    function renderTable() {
        giroTableBody.innerHTML = '';

        if (filteredData.length === 0) {
            giroTableBody.innerHTML = `
                <tr>
                    <td colspan="8" style="text-align: center; padding: 2rem; color: var(--text-muted); font-style: italic;">
                        Nenhum item corresponde aos filtros e buscas ativos.
                    </td>
                </tr>
            `;
            return;
        }

        filteredData.forEach(item => {
            const tr = document.createElement('tr');
            
            // Add custom visual cue classes based on purchase status
            tr.className = item.comprado ? 'repost-row' : 'no-repost-row';

            const badgeHtml = item.comprado 
                ? '<span class="badge-sim">🟢 Sim</span>' 
                : '<span class="badge-nao">🔴 Não</span>';

            const recurHtml = getRecorrenciaBadge(item.recorrencia);

            tr.innerHTML = `
                <td style="font-weight: 700; color: #fff;">${item.code}</td>
                <td>
                    <div style="font-weight: 500;">${item.desc}</div>
                </td>
                <td style="text-align: center; color: var(--text-muted); font-size: 0.8rem;">${item.un}</td>
                <td style="text-align: center;">${badgeHtml}</td>
                <td style="text-align: center;">${recurHtml}</td>
                <td style="text-align: center; font-weight: 700; color: ${item.comprado ? 'var(--success)' : 'var(--text-muted)'}">${item.qty.toLocaleString('pt-BR')}</td>
                <td style="color: #cbd5e1; max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${item.fornecedor}</td>
                <td style="text-align: right; font-weight: 700; color: ${item.comprado ? '#34d399' : '#f43f5e'}">${formatBRL(item.valor)}</td>
            `;

            giroTableBody.appendChild(tr);
        });
    }

    // --- Chart Implementations & Live Updates ---
    function initCharts() {
        const primaryFont = { family: 'Outfit, sans-serif' };

        // 1. Doughnut Chart: Sim vs Não
        const doughnutCtx = document.getElementById('replenish-doughnut').getContext('2d');
        doughnutChart = new Chart(doughnutCtx, {
            type: 'doughnut',
            data: {
                labels: ['Repostos (Sim)', 'Não Repostos (Não)'],
                datasets: [{
                    data: [0, 0],
                    backgroundColor: ['#10b981', '#f43f5e'],
                    borderColor: '#0f172a',
                    borderWidth: 2,
                    hoverOffset: 12
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '72%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { color: '#94a3b8', font: { family: 'Outfit', size: 11 } }
                    },
                    datalabels: {
                        color: '#fff',
                        backgroundColor: 'rgba(15, 23, 42, 0.85)',
                        borderRadius: 4,
                        padding: 6,
                        font: { weight: 'bold', size: 10 },
                        formatter: (val, ctx) => {
                            let sum = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                            return sum > 0 ? `${(val * 100 / sum).toFixed(0)}%` : '';
                        }
                    }
                }
            },
            plugins: [ChartDataLabels]
        });

        // 2. Horizontal Bar: Supplier Spend
        const supplierCtx = document.getElementById('suppliers-value-bar').getContext('2d');
        supplierChart = new Chart(supplierCtx, {
            type: 'bar',
            data: {
                labels: [],
                datasets: [{
                    label: 'Valor Reposto (R$)',
                    data: [],
                    backgroundColor: 'rgba(59, 130, 246, 0.85)',
                    borderColor: '#2563eb',
                    borderWidth: 1.5,
                    borderRadius: 4
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    datalabels: { display: false }
                },
                scales: {
                    x: {
                        grid: { color: 'rgba(255, 255, 255, 0.05)' },
                        ticks: { color: '#94a3b8', font: { size: 9 } }
                    },
                    y: {
                        grid: { display: false },
                        ticks: { color: '#f1f5f9', font: { family: 'Inter', size: 9 } }
                    }
                }
            }
        });

        // 3. Vertical Bar: Top items spent
        const itemCtx = document.getElementById('items-value-bar').getContext('2d');
        itemChart = new Chart(itemCtx, {
            type: 'bar',
            data: {
                labels: [],
                datasets: [{
                    label: 'Valor Gasto (R$)',
                    data: [],
                    backgroundColor: 'rgba(245, 158, 11, 0.85)',
                    borderColor: '#f59e0b',
                    borderWidth: 1.5,
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    datalabels: { display: false }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { color: '#f1f5f9', font: { size: 9 }, maxRotation: 45, minRotation: 45 }
                    },
                    y: {
                        grid: { color: 'rgba(255, 255, 255, 0.05)' },
                        ticks: { color: '#94a3b8', font: { size: 9 } }
                    }
                }
            }
        });

        updateChartsData();
    }

    function updateChartsData() {
        if (!doughnutChart) return;

        // --- Doughnut Update ---
        const repCount = filteredData.filter(i => i.comprado).length;
        const nonRepCount = filteredData.filter(i => !i.comprado).length;
        doughnutChart.data.datasets[0].data = [repCount, nonRepCount];
        doughnutChart.update();

        // --- Supplier Bar Update ---
        const suppSpend = {};
        filteredData.filter(i => i.comprado).forEach(item => {
            suppSpend[item.fornecedor] = (suppSpend[item.fornecedor] || 0) + item.valor;
        });
        const topSuppliers = Object.entries(suppSpend)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10);

        supplierChart.data.labels = topSuppliers.map(s => s[0].length > 25 ? s[0].substring(0, 22) + '...' : s[0]);
        supplierChart.data.datasets[0].data = topSuppliers.map(s => s[1]);
        supplierChart.update();

        // --- Items Bar Update ---
        const topItems = filteredData.filter(i => i.comprado)
            .sort((a, b) => b.valor - a.valor)
            .slice(0, 10);

        itemChart.data.labels = topItems.map(i => i.code);
        itemChart.data.datasets[0].data = topItems.map(i => i.valor);
        
        // Add descriptions inside tooltips easily
        itemChart.options.plugins.tooltip = {
            callbacks: {
                title: (context) => {
                    const idx = context[0].dataIndex;
                    return `Cód: ${topItems[idx].code} - ${topItems[idx].desc}`;
                }
            }
        };
        
        itemChart.update();
    }

    // --- Search & Search Events ---
    simpleSearch.addEventListener('input', applyFiltersAndSort);
    advancedSearch.addEventListener('input', applyFiltersAndSort);

    // --- Filter Buttons Events ---
    filterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            filterBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            activeFilter = btn.getAttribute('data-filter');
            applyFiltersAndSort();
        });
    });

    // --- Clear Filters ---
    if (clearFiltersBtn) {
        clearFiltersBtn.addEventListener('click', () => {
            simpleSearch.value = '';
            advancedSearch.value = '';
            activeFilter = 'all';
            sortColumn = 'none';
            sortDirection = 'none';

            filterBtns.forEach(b => {
                if (b.getAttribute('data-filter') === 'all') b.classList.add('active');
                else b.classList.remove('active');
            });

            resetSortIcons();
            applyFiltersAndSort();
        });
    }

    // --- Sorting Interaction ---
    function resetSortIcons(...excludes) {
        const headers = [sortCodeHeader, sortDescHeader, sortQtyHeader, sortValueHeader, sortRecorrenciaHeader];
        headers.forEach(h => {
            if (h && !excludes.includes(h)) {
                h.querySelector('.sort-icon').textContent = '↕️';
            }
        });
    }

    function toggleSort(columnName, headerEl) {
        if (sortColumn === columnName) {
            sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            sortColumn = columnName;
            sortDirection = 'asc';
        }

        resetSortIcons(headerEl);
        headerEl.querySelector('.sort-icon').textContent = sortDirection === 'asc' ? '🔼' : '🔽';
        applyFiltersAndSort();
    }

    if (sortCodeHeader) sortCodeHeader.addEventListener('click', () => toggleSort('code', sortCodeHeader));
    if (sortDescHeader) sortDescHeader.addEventListener('click', () => toggleSort('desc', sortDescHeader));
    if (sortRecorrenciaHeader) sortRecorrenciaHeader.addEventListener('click', () => toggleSort('recorrencia', sortRecorrenciaHeader));
    if (sortQtyHeader) sortQtyHeader.addEventListener('click', () => toggleSort('qty', sortQtyHeader));
    if (sortValueHeader) sortValueHeader.addEventListener('click', () => toggleSort('value', sortValueHeader));

    // --- Excel Export ---
    if (exportExcelBtn) {
        exportExcelBtn.addEventListener('click', () => {
            if (filteredData.length === 0) {
                alert('Não há dados ativos para exportar.');
                return;
            }

            const exportRows = filteredData.map(i => ({
                "Código do produto": i.code,
                "Descrição do produto": i.desc,
                "Unidade de medida": i.un,
                "Comprado nos últimos 6 meses": i.comprado ? 'Sim' : 'Não',
                "Recorrência (%)": i.recorrencia !== undefined ? `${i.recorrencia.toFixed(1)}%` : '0%',
                "Quantidade comprada": i.qty,
                "Fornecedor": i.fornecedor,
                "Valor gasto": i.valor
            }));

            const ws = XLSX.utils.json_to_sheet(exportRows);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Levantamento Giro Filtrado");

            // Beautify widths
            ws['!cols'] = [
                { wch: 15 }, // Codigo
                { wch: 45 }, // Descricao
                { wch: 10 }, // UN
                { wch: 25 }, // Comprado
                { wch: 18 }, // Recorrência (%)
                { wch: 18 }, // Qtd
                { wch: 35 }, // Fornecedor
                { wch: 15 }  // Valor
            ];

            const today = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
            XLSX.writeFile(wb, `levantamento_giro_filtrado_${today}.xlsx`);
        });
    }

});
