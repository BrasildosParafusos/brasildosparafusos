/**
 * Inteligência Evolutiva de Abastecimento v5.0
 * Módulo de Análise Histórica e Evolução Temporal de Criticidade de Estoque
 */
document.addEventListener('DOMContentLoaded', () => {
    console.log("Inteligência Evolutiva v5.0 Inicializada");

    const CRITICIDADE = {
        'rupture': { label: 'Ruptura', weight: 3, color: '#fb7185' },
        'attention': { label: 'Atenção', weight: 2, color: '#f59e0b' },
        'suggest': { label: 'Sugestão', weight: 1, color: '#818cf8' },
        'ok': { label: 'Seguro', weight: 0, color: '#34d399' },
        'ignored': { label: 'Ignorado', weight: 0, color: '#9ca3af' }
    };

    const RECORRENCIA_MINIMA = 0.17;

    let snapshotHistory = [];
    let baseSnapshot = null;
    let currentSnapshot = null;
    let currentTimelineIdx = 0;
    let activeBuyer = 'all';
    let activeSupplier = 'all';
    let evolutionChart = null;
    let buyerMap = JSON.parse(localStorage.getItem('buyerMap') || '{}');
    
    let comparisonMode = 'historica'; // 'historica' ou 'diaria'
    let itemStatsMap = new Map(); // Mapa global de histórico de cada item (código -> stats)

    const folderInputs = [document.getElementById('folder-upload'), document.getElementById('folder-upload-welcome')];
    const mainContent = document.getElementById('main-content');
    const welcomeState = document.getElementById('welcome-state');
    const loadingOverlay = document.getElementById('loading-overlay');
    const loadingText = loadingOverlay.querySelector('.loading-text');
    const timelineRange = document.getElementById('timeline-range');
    const timelineTicks = document.getElementById('timeline-ticks');
    const tableBody = document.getElementById('evolution-table-body');
    const tableSearch = document.getElementById('table-search');
    const evolutionFilter = document.getElementById('evolution-filter');

    // Botões de Comparação
    const btnCompHistorica = document.getElementById('btn-comp-historica');
    const btnCompDiaria = document.getElementById('btn-comp-diaria');

    if (btnCompHistorica && btnCompDiaria) {
        btnCompHistorica.onclick = () => {
            comparisonMode = 'historica';
            btnCompHistorica.classList.add('active');
            btnCompDiaria.classList.remove('active');
            updateDashboard();
        };

        btnCompDiaria.onclick = () => {
            comparisonMode = 'diaria';
            btnCompDiaria.classList.add('active');
            btnCompHistorica.classList.remove('active');
            updateDashboard();
        };
    }

    function parseNumeric(val) {
        if (val === undefined || val === null || val === '') return 0;
        if (typeof val === 'number') return val;
        let str = val.toString().replace('R$', '').replace(/\s/g, '').trim();
        if (str.startsWith('.') || str.startsWith(',')) str = '0' + str;
        const hasComma = str.includes(',');
        const hasDot = str.includes('.');
        if (hasComma && hasDot) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) str = str.replace(/\./g, '').replace(',', '.');
            else str = str.replace(/,/g, '');
        } else if (hasComma) {
            const parts = str.split(',');
            if (parts.length > 2 || parts[1].length === 3) str = str.replace(/,/g, '');
            else str = str.replace(',', '.');
        }
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }

    function findColumn(headers, aliases) {
        const cleanHeaders = headers.map(h => String(h || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        const cleanAliases = aliases.map(a => a.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        for (let alias of cleanAliases) {
            const idx = cleanHeaders.indexOf(alias);
            if (idx !== -1) return headers[idx];
        }
        for (let alias of cleanAliases) {
            const idx = cleanHeaders.findIndex(h => h.includes(alias));
            if (idx !== -1) return headers[idx];
        }
        return null;
    }

    async function processExcelFile(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    let headerIndex = -1;
                    for (let i = 0; i < Math.min(rawRows.length, 30); i++) {
                        if (rawRows[i] && rawRows[i].some(c => {
                            const s = String(c||'').toLowerCase();
                            return s.includes('estoque') || s.includes('produto') || s.includes('descri');
                        })) {
                            headerIndex = i;
                            break;
                        }
                    }
                    if (headerIndex === -1) return resolve(null);

                    const rawHeaders = rawRows[headerIndex];
                    
                    // Normalizar cabeçalhos para strings limpas, tratando datas vindas do XLSX
                    const headers = rawHeaders.map(h => {
                        if (h instanceof Date) {
                            const m = (h.getMonth() + 1).toString().padStart(2, '0');
                            const y = h.getFullYear();
                            return `${m}/${y}`;
                        }
                        return String(h || '').trim();
                    });

                    function findColumnIndex(headersList, aliases) {
                        const cleanHeaders = headersList.map(h => String(h || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
                        const cleanAliases = aliases.map(a => a.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
                        for (let alias of cleanAliases) {
                            const idx = cleanHeaders.indexOf(alias);
                            if (idx !== -1) return idx;
                        }
                        for (let alias of cleanAliases) {
                            const idx = cleanHeaders.findIndex(h => h.includes(alias));
                            if (idx !== -1) return idx;
                        }
                        return -1;
                    }

                    const colIdx = {
                        code: findColumnIndex(headers, ['Produto', 'Código', 'Item', 'Cód.', 'ID']),
                        desc: findColumnIndex(headers, ['Descrição longa do produto', 'Descrição', 'Desc', 'Nome', 'Produto Descrição', 'Texto']),
                        estoque: findColumnIndex(headers, ['Estoque', 'Saldo', 'Qtd. Estoque', 'Estoque Total', 'Saldo Atual', 'Saldo Disponível', 'Disp.', 'Qtd. Disponível', 'Estoque Atual']),
                        encomendas: findColumnIndex(headers, [
                            'Encomendas', 'Qtd. Encomenda', 'Saldo Pedido Compra', 'Saldo Ped. Compra', 'Pedido Compra', 
                            'Qtd. em Pedido Compra', 'Qtd. no Pedido Compra', 'Saldo a Receber', 'A Receber', 'Pedidos', 
                            'Qtd. Pedida', 'Saldo Pedido', 'Compras', 'Qtd em Pedido', 'Qtd. Ped.', 'Saldo Ped.', 
                            'Pendência', 'Qtd. no Pedido', 'Encomenda', 'Pedido', 'Qtd Ped Compra', 'A Receber Total',
                            'A Entregar', 'Saldo a Entregar', 'Qtd. Pendente', 'Pendente', 'Saldo O.C.', 'Ord. Compra'
                        ]),
                        medVenda: findColumnIndex(headers, ['Méd.Venda', 'Média Venda', 'Venda Média', 'Saída Média', 'Média']),
                        fornecedor: findColumnIndex(headers, ['Razão social do fornecedor', 'Fornecedor', 'Fornec', 'Fabricante', 'Último Fornecedor', 'Fornecedor Principal', 'Nome Fornecedor']),
                        comprador: findColumnIndex(headers, ['Comprador', 'Responsável', 'Buyer']),
                        totalSales: findColumnIndex(headers, ['Vendas Total', 'Total Vendas', 'Saída Total', 'Vendas', 'Qtd. Vendida', 'Venda Total', 'Venda', 'Saídas', 'Giro']),
                        custo: findColumnIndex(headers, ['Preço reposição', 'Custo aquisição', 'Custo Unitário', 'Custo', 'Preço Custo', 'Vlr. Custo', 'Custo Médio', 'Unitário'])
                    };

                    // Filtrar colunas de meses (formato MM/YYYY ou M/YYYY) - Excluir meses futuros após 05/2026
                    const monthMap = [];
                    headers.forEach((h, idx) => {
                        if (h && /^\d{1,2}\/\d{2,4}$/.test(h)) {
                            const [mon, yr] = h.split('/').map(Number);
                            const year = yr < 100 ? 2000 + yr : yr;
                            if (year < 2026 || (year === 2026 && mon <= 5)) {
                                monthMap.push({ index: idx, label: h });
                            }
                        }
                    });

                    // Ordenar os meses cronologicamente para consistência
                    monthMap.sort((a, b) => {
                        const [mA, yA] = a.label.split('/').map(Number);
                        const [mB, yB] = b.label.split('/').map(Number);
                        const yearA = yA < 100 ? 2000 + yA : yA;
                        const yearB = yB < 100 ? 2000 + yB : yB;
                        return (yearA * 12 + mA) - (yearB * 12 + mB);
                    });

                    const dataRows = rawRows.slice(headerIndex + 1);
                    const itemsMap = new Map();

                    dataRows.forEach(row => {
                        if (!row || row.length === 0) return;

                        const codeVal = colIdx.code !== -1 ? row[colIdx.code] : null;
                        const code = String(codeVal || '').trim();
                        if (!code || code === 'undefined' || code === 'null') return;

                        const estoque = parseNumeric(colIdx.estoque !== -1 ? row[colIdx.estoque] : 0);
                        const encomendas = parseNumeric(colIdx.encomendas !== -1 ? row[colIdx.encomendas] : 0);
                        const salesTotal = parseNumeric(colIdx.totalSales !== -1 ? row[colIdx.totalSales] : 0);
                        let medVenda = parseNumeric(colIdx.medVenda !== -1 ? row[colIdx.medVenda] : 0);
                        let custoUnitario = parseNumeric(colIdx.custo !== -1 ? row[colIdx.custo] : 1);
                        let recorrencia = 0;

                        if (monthMap.length > 0) {
                            let activeMonths = 0;
                            let sum = 0;
                            monthMap.forEach(m => {
                                const v = parseNumeric(row[m.index]);
                                if (v > 0) { activeMonths++; sum += v; }
                            });
                            recorrencia = activeMonths / monthMap.length;
                            if (activeMonths > 0) medVenda = sum / activeMonths;
                        }

                        let status = 'ok';
                        const disponivel = estoque + encomendas;
                        if (recorrencia > RECORRENCIA_MINIMA) {
                            if (medVenda > disponivel) status = 'rupture';
                            else if ((medVenda * 2) > disponivel) status = 'attention';
                            else if ((medVenda * 3) > disponivel) status = 'suggest';
                        } else if (salesTotal === 0 && monthMap.length > 0) {
                            status = 'ignored';
                        }

                        let comprador = colIdx.comprador !== -1 ? row[colIdx.comprador] : null;
                        if (!comprador) {
                            const cleanCode = code.replace(/^0+/, '').replace(/[.]/g, '');
                            comprador = buyerMap[code] || buyerMap[cleanCode] || 'N/D';
                        }

                        const descVal = colIdx.desc !== -1 ? row[colIdx.desc] : 'S/D';
                        const fornecedorVal = colIdx.fornecedor !== -1 ? row[colIdx.fornecedor] : 'N/D';

                        itemsMap.set(code, {
                            code,
                            desc: descVal || 'S/D',
                            fornecedor: fornecedorVal || 'N/D',
                            comprador: comprador || 'N/D',
                            status,
                            weight: CRITICIDADE[status].weight,
                            value: medVenda * custoUnitario,
                            estoque: disponivel,
                            venda: medVenda
                        });
                    });

                    let fileDate = file.name.match(/(\d{2})[.\/](\d{2})[.\/](\d{4})/);
                    let dateStr = fileDate ? `${fileDate[3]}-${fileDate[2]}-${fileDate[1]}` : new Date().toISOString().split('T')[0];

                    resolve({ name: file.name, date: dateStr, itemsMap });
                } catch (err) {
                    console.error("Erro ao processar arquivo:", file.name, err);
                    resolve(null);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    async function handleFiles(files) {
        try {
            loadingOverlay.style.display = 'flex';
            const filteredFiles = Array.from(files).filter(f => f.name.match(/\.(xlsx|xls)$/i) && !f.name.startsWith('~$'));
            
            if (filteredFiles.length === 0) {
                loadingOverlay.style.display = 'none';
                alert("Nenhum arquivo Excel encontrado na pasta selecionada.");
                return;
            }

            const snaps = [];
            let processedCount = 0;

            for (let f of filteredFiles) {
                processedCount++;
                loadingText.innerHTML = `Processando arquivos... (${processedCount}/${filteredFiles.length})<br><small style="opacity:0.7">${f.name}</small>`;
                const res = await processExcelFile(f);
                if (res) snaps.push(res);
            }

            if (snaps.length === 0) {
                loadingOverlay.style.display = 'none';
                alert("Não foi possível extrair dados dos arquivos. Verifique os cabeçalhos.");
                return;
            }

            // Ordenação estrita por data cronológica
            snaps.sort((a, b) => new Date(a.date) - new Date(b.date));
            snapshotHistory = snaps;
            
            // O primeiro arquivo da pasta em ordem cronológica é a base histórica
            baseSnapshot = snaps[0];
            currentTimelineIdx = snaps.length - 1;

            // Consolidar e calcular métricas históricas gerais por item
            calculateGlobalItemStats();

            updateDashboard();
            
            welcomeState.style.display = 'none';
            mainContent.style.display = 'block';
            loadingOverlay.style.display = 'none';

        } catch (error) {
            console.error("Erro crítico no handleFiles:", error);
            loadingOverlay.style.display = 'none';
            alert("Ocorreu um erro ao processar os dados: " + error.message);
        }
    }

    /**
     * Calcula as métricas históricas de cada produto ao longo de TODA a linha temporal.
     */
    function calculateGlobalItemStats() {
        itemStatsMap.clear();
        
        // Obter todos os códigos únicos de produtos de todo o histórico
        const allCodes = new Set();
        snapshotHistory.forEach(snap => {
            snap.itemsMap.forEach((item, code) => {
                allCodes.add(code);
            });
        });

        allCodes.forEach(code => {
            let stats = {
                code: code,
                desc: 'S/D',
                fornecedor: 'N/D',
                comprador: 'N/D',
                history: [], // Histórico cronológico
                ruptureCount: 0,
                attentionCount: 0,
                suggestCount: 0,
                okCount: 0,
                criticalDays: 0, // snaps em que esteve crítico (weight > 0)
                recuperacoesCount: 0, // quantidade de vezes que a criticidade diminuiu entre consecutivos
                worsenedCount: 0, // quantidade de vezes que a criticidade aumentou entre consecutivos
                reincidente: false // se saiu de crítico -> ok -> e voltou a ser crítico
            };

            let wasCritical = false;
            let wasRecovered = false;
            let lastWeight = null;

            snapshotHistory.forEach(snap => {
                const item = snap.itemsMap.get(code);
                const status = item ? item.status : 'ok';
                const weight = item ? item.weight : 0;
                
                if (item) {
                    if (item.desc && item.desc !== 'S/D') stats.desc = item.desc;
                    if (item.fornecedor && item.fornecedor !== 'N/D') stats.fornecedor = item.fornecedor;
                    if (item.comprador && item.comprador !== 'N/D') stats.comprador = item.comprador;
                }

                const isMissing = !item;
                stats.history.push({ date: snap.date, status, weight, missing: isMissing });

                // Acumulador de criticidade
                if (status === 'rupture') stats.ruptureCount++;
                else if (status === 'attention') stats.attentionCount++;
                else if (status === 'suggest') stats.suggestCount++;
                else if (status === 'ok') stats.okCount++;

                if (weight > 0) {
                    stats.criticalDays++;
                }

                // Transições consecutivas
                if (lastWeight !== null) {
                    const diff = lastWeight - weight;
                    if (diff > 0) {
                        stats.recuperacoesCount++;
                    } else if (diff < 0) {
                        stats.worsenedCount++;
                    }
                }

                // Regra de Reincidência (Crítico -> OK real -> Crítico)
                if (!isMissing) {
                    if (weight > 0) {
                        if (!wasCritical && wasRecovered) {
                            stats.reincidente = true;
                        }
                        wasCritical = true;
                    } else if (weight === 0 && wasCritical) {
                        wasCritical = false;
                        wasRecovered = true;
                    }
                }

                lastWeight = weight;
            });

            itemStatsMap.set(code, stats);
        });
    }

    function updateDashboard() {
        try {
            const snap = snapshotHistory[currentTimelineIdx];
            if (!snap) return;

            currentSnapshot = snap;
            document.getElementById('base-date-display').textContent = baseSnapshot.date.split('-').reverse().join('/');
            document.getElementById('current-snapshot-date').textContent = snap.date.split('-').reverse().join('/');

            renderTimeline();
            calculateEvolution();
            renderCharts();
            renderTable();
            updatePerformance();
            updateBuyerFilter();
        } catch (e) {
            console.error("Erro ao atualizar dashboard:", e);
            alert("Erro ao renderizar dados: " + e.message);
            loadingOverlay.style.display = 'none';
        }
    }

    function renderTimeline() {
        timelineRange.max = snapshotHistory.length - 1;
        timelineRange.value = currentTimelineIdx;
        
        timelineTicks.innerHTML = snapshotHistory.map((s, idx) => `
            <span class="${idx === currentTimelineIdx ? 'active' : ''}" onclick="window.jumpTo(${idx})">
                ${s.date.split('-').slice(1).reverse().join('/')}
            </span>
        `).join('');
    }

    window.jumpTo = (idx) => {
        currentTimelineIdx = idx;
        updateDashboard();
    };

    function calculateEvolution() {
        if (snapshotHistory.length === 0) return;

        // Definir instantâneo de referência baseado no modo de comparação
        let refSnapshot = baseSnapshot;
        if (comparisonMode === 'diaria') {
            refSnapshot = currentTimelineIdx > 0 ? snapshotHistory[currentTimelineIdx - 1] : currentSnapshot;
        }

        let recovered = { rupture: 0, attention: 0, suggest: 0, total: 0 };
        let worsened = 0;
        let stable = 0;
        let totalRefCritical = 0;
        let refRuptures = 0;
        let refAttentions = 0;
        let refSuggests = 0;
        let newItems = 0;
        let discontinuedItems = 0;

        Array.from(itemStatsMap.values())
            .filter(stats => (activeBuyer === 'all' || stats.comprador === activeBuyer) && (activeSupplier === 'all' || stats.fornecedor === activeSupplier))
            .forEach(stats => {
                const refItem = refSnapshot.itemsMap.get(stats.code);
                const currentItem = currentSnapshot.itemsMap.get(stats.code);
                
                if (!refItem && currentItem) newItems++;
                if (refItem && !currentItem) discontinuedItems++;

                // Lógica de Piora e Estabilidade (Retrato Atual)
                if (refItem && currentItem) {
                    const evolution = refItem.weight - currentItem.weight;
                    if (evolution < 0) {
                        worsened++;
                    } else if (evolution === 0) {
                        stable++;
                    }
                }

                // Lógica Histórica Acumulada para Recuperações e Reincidência
                // A base é sempre o baseSnapshot (início do histórico)
                const baseItem = baseSnapshot.itemsMap.get(stats.code);
                const baseWeight = baseItem ? baseItem.weight : 0;
                const baseStatus = baseItem ? baseItem.status : 'ok';
                
                let isRecurrent = false;
                let histRecovered = false;

                if (baseWeight > 0) {
                    // Contabilizar no denominador de críticos da base
                    totalRefCritical++;
                    if (baseStatus === 'rupture') refRuptures++;
                    else if (baseStatus === 'attention') refAttentions++;
                    else if (baseStatus === 'suggest') refSuggests++;

                    let minWeightSeen = baseWeight;

                    // Percorrer a timeline até a data atual para avaliar recuperação acumulada
                    for (let i = 1; i <= currentTimelineIdx; i++) {
                        const snap = snapshotHistory[i];
                        const itemAtTime = snap.itemsMap.get(stats.code);
                        if (itemAtTime) {
                            if (itemAtTime.weight < baseWeight) {
                                histRecovered = true;
                            }
                            if (itemAtTime.weight < minWeightSeen) {
                                minWeightSeen = itemAtTime.weight;
                            }
                            if (histRecovered && itemAtTime.weight > minWeightSeen) {
                                isRecurrent = true;
                            }
                        }
                    }

                    if (histRecovered) {
                        recovered.total++;
                        if (baseStatus === 'rupture') recovered.rupture++;
                        else if (baseStatus === 'attention') recovered.attention++;
                        else if (baseStatus === 'suggest') recovered.suggest++;
                    }
                }
                
                stats.isRecurrentAtCurrentIdx = isRecurrent;
                stats.histRecoveredAtCurrentIdx = histRecovered;
            });

        // Atualizar os KPIs
        document.getElementById('global-recovered-count').textContent = recovered.total;
        document.getElementById('rupture-recovered').textContent = recovered.rupture;
        document.getElementById('attention-recovered').textContent = recovered.attention;
        document.getElementById('suggest-recovered').textContent = recovered.suggest;
        document.getElementById('worsened-count').textContent = worsened;
        document.getElementById('stable-count').textContent = stable;
        
        const newItemsEl = document.getElementById('new-items-count');
        if (newItemsEl) newItemsEl.textContent = newItems;
        const discontinuedEl = document.getElementById('discontinued-count');
        if (discontinuedEl) discontinuedEl.textContent = discontinuedItems;

        document.getElementById('rupture-recovered-percent').textContent = refRuptures > 0 ? `${(recovered.rupture / refRuptures * 100).toFixed(1)}% de sucesso | ${refRuptures - recovered.rupture} não recuperadas` : '0% de sucesso';
        document.getElementById('attention-recovered-percent').textContent = refAttentions > 0 ? `${(recovered.attention / refAttentions * 100).toFixed(1)}% de sucesso | ${refAttentions - recovered.attention} não recuperadas` : '0% de sucesso';
        document.getElementById('suggest-recovered-percent').textContent = refSuggests > 0 ? `${(recovered.suggest / refSuggests * 100).toFixed(1)}% de sucesso | ${refSuggests - recovered.suggest} não recuperadas` : '0% de sucesso';

        const efficiency = totalRefCritical > 0 ? (recovered.total / totalRefCritical * 100) : 0;
        document.getElementById('efficiency-percent').textContent = `${efficiency.toFixed(1)}%`;

        if (comparisonMode === 'diaria') {
            document.getElementById('global-recovered-sub').textContent = `Representa ${efficiency.toFixed(1)}% dos ${totalRefCritical} itens em nível de ruptura no dia anterior (${totalRefCritical - recovered.total} não recuperados)`;
        } else {
            document.getElementById('global-recovered-sub').textContent = `Representa ${efficiency.toFixed(1)}% dos ${totalRefCritical} itens em nível de ruptura na base (${totalRefCritical - recovered.total} não recuperados)`;
        }

        // Indicadores Históricos Globais de Abastecimento (Acumulados da Linha do Tempo)
        let relapsedTotal = 0;

        itemStatsMap.forEach(stats => {
            if (activeBuyer === 'all' || stats.comprador === activeBuyer) {
                if (stats.isRecurrentAtCurrentIdx) {
                    relapsedTotal++;
                }
            }
        });

        document.getElementById('relapsed-count').textContent = relapsedTotal;

        // Sincronizar classes de destaque dos cards baseados no filtro atual
        const currentVal = evolutionFilter.value;
        const cardFilters = {
            'card-recovery': 'improved',
            'card-rupture-recovered': 'improved_rupture',
            'card-attention-recovered': 'improved_attention',
            'card-suggest-recovered': 'improved_suggest',
            'card-relapsed': 'relapsed',
            'card-stable': 'stable',
            'card-worsened': 'worsened',
            'card-new-items': 'new_items',
            'card-discontinued': 'discontinued'
        };
        Object.entries(cardFilters).forEach(([cardId, filterVal]) => {
            const el = document.getElementById(cardId);
            if (el) {
                if (filterVal === currentVal) {
                    el.classList.add('selected');
                } else {
                    el.classList.remove('selected');
                }
            }
        });
    }

    function renderCharts() {
        const canvas = document.getElementById('evolution-chart');
        if (!canvas) return;
        const ctx = canvas.getContext('2d');
        if (evolutionChart) evolutionChart.destroy();

        const historyLabels = snapshotHistory.map(s => s.date.split('-').reverse().slice(0,2).join('/'));
        
        const filterVal = evolutionFilter ? evolutionFilter.value : 'all';
        let filteredCodes = null;
        if (filterVal !== 'all') {
            let refSnapshot = baseSnapshot;
            if (comparisonMode === 'diaria') {
                refSnapshot = currentTimelineIdx > 0 ? snapshotHistory[currentTimelineIdx - 1] : currentSnapshot;
            }
            
            filteredCodes = new Set(
                Array.from(itemStatsMap.values())
                    .filter(stats => (activeBuyer === 'all' || stats.comprador === activeBuyer) && (activeSupplier === 'all' || stats.fornecedor === activeSupplier))
                    .map(stats => {
                        const refItem = refSnapshot.itemsMap.get(stats.code);
                        const itemRef = refItem || { status: 'ok', weight: 0, desc: stats.desc, comprador: stats.comprador, fornecedor: stats.fornecedor };

                        const currentItem = currentSnapshot.itemsMap.get(stats.code);
                        const itemCurrent = currentItem || { status: 'ok', weight: 0, desc: stats.desc, comprador: stats.comprador, fornecedor: stats.fornecedor };

                        const evolution = itemRef.weight - itemCurrent.weight;
                        const isNew = !refItem && currentItem;
                        const isDiscontinued = refItem && !currentItem;
                        return { code: stats.code, itemRef, itemCurrent, evolution, stats, isNew, isDiscontinued };
                    })
                    .filter(row => {
                        if (filterVal === 'improved') return row.evolution > 0;
                        if (filterVal === 'improved_rupture') return row.evolution > 0 && row.itemRef.status === 'rupture';
                        if (filterVal === 'improved_attention') return row.evolution > 0 && row.itemRef.status === 'attention';
                        if (filterVal === 'improved_suggest') return row.evolution > 0 && row.itemRef.status === 'suggest';
                        if (filterVal === 'worsened') return row.evolution < 0;
                        if (filterVal === 'stable') return row.evolution === 0;
                        if (filterVal === 'relapsed') return row.stats.reincidente === true;
                        if (filterVal === 'new_items') return row.isNew;
                        if (filterVal === 'discontinued') return row.isDiscontinued;
                        return true;
                    })
                    .map(r => r.code)
            );
        }

        // Contar quantos itens estão em algum nível de ruptura (rupture, attention, suggest)
        const criticalHistory = snapshotHistory.map(snap => {
            const items = Array.from(snap.itemsMap.values())
                .filter(i => (activeBuyer === 'all' || i.comprador === activeBuyer) && 
                             (activeSupplier === 'all' || i.fornecedor === activeSupplier) &&
                             (filteredCodes === null || filteredCodes.has(i.code)));
            return items.filter(i => i.status === 'rupture' || i.status === 'attention' || i.status === 'suggest').length;
        });

        // Contar quantos itens estão seguros (status === 'ok')
        const safeHistory = snapshotHistory.map(snap => {
            const items = Array.from(snap.itemsMap.values())
                .filter(i => (activeBuyer === 'all' || i.comprador === activeBuyer) && 
                             (activeSupplier === 'all' || i.fornecedor === activeSupplier) &&
                             (filteredCodes === null || filteredCodes.has(i.code)));
            return items.filter(i => i.status === 'ok').length;
        });

        evolutionChart = new Chart(ctx, {
            plugins: [ChartDataLabels],
            type: 'line',
            data: {
                labels: historyLabels,
                datasets: [
                    {
                        label: 'Itens Seguros',
                        data: safeHistory,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.08)',
                        fill: true,
                        tension: 0.4,
                        pointRadius: snapshotHistory.length > 30 ? 2 : 5,
                        datalabels: {
                            color: '#34d399',
                            align: 'top',
                            anchor: 'end',
                            offset: 6,
                            textShadowColor: 'rgba(0, 0, 0, 0.8)',
                            textShadowBlur: 4
                        }
                    },
                    {
                        label: 'Itens em Criticidade (Rupturas)',
                        data: criticalHistory,
                        borderColor: '#fb7185',
                        backgroundColor: 'rgba(251, 113, 133, 0.08)',
                        fill: true,
                        tension: 0.4,
                        pointRadius: snapshotHistory.length > 30 ? 2 : 5,
                        datalabels: {
                            color: '#fb7185',
                            align: 'bottom',
                            anchor: 'start',
                            offset: 6,
                            textShadowColor: 'rgba(0, 0, 0, 0.8)',
                            textShadowBlur: 4
                        }
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        top: 35,
                        bottom: 10,
                        left: 10,
                        right: 180
                    }
                },
                plugins: { 
                    legend: { 
                        display: false
                    },
                    datalabels: {
                        display: true,
                        font: {
                            family: 'Outfit',
                            size: 11,
                            weight: '800'
                        },
                        formatter: (val, ctx) => {
                            const idx = ctx.dataIndex;
                            const data = ctx.dataset.data;
                            let arrow = '';
                            if (idx > 0) {
                                const prev = data[idx - 1];
                                if (val > prev) {
                                    arrow = '↑ ';
                                } else if (val < prev) {
                                    arrow = '↓ ';
                                }
                            }
                            const mainText = arrow + val.toLocaleString('pt-BR');
                            if (idx === data.length - 1) {
                                return [mainText, ctx.dataset.label];
                            }
                            return mainText;
                        }
                    }
                },
                scales: {
                    y: { 
                        beginAtZero: true, 
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        ticks: { color: '#94a3b8' }
                    },
                    x: { 
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    }
                }
            }
        });
    }

    function renderTable() {
        const searchTerm = tableSearch.value.toLowerCase();
        const filterVal = evolutionFilter.value;
        
        let refSnapshot = baseSnapshot;
        if (comparisonMode === 'diaria') {
            refSnapshot = currentTimelineIdx > 0 ? snapshotHistory[currentTimelineIdx - 1] : currentSnapshot;
        }

        const rows = Array.from(itemStatsMap.values())
            .filter(stats => (activeBuyer === 'all' || stats.comprador === activeBuyer) && (activeSupplier === 'all' || stats.fornecedor === activeSupplier))
            .map(stats => {
                const refItem = refSnapshot.itemsMap.get(stats.code);
                const itemRef = refItem || { status: 'ok', weight: 0, desc: stats.desc, comprador: stats.comprador, fornecedor: stats.fornecedor, estoque: 0, venda: 0 };

                const baseItem = baseSnapshot.itemsMap.get(stats.code) || { status: 'ok', weight: 0 };

                const currentItem = currentSnapshot.itemsMap.get(stats.code);
                const itemCurrent = currentItem || { status: 'ok', weight: 0, desc: stats.desc, comprador: stats.comprador, fornecedor: stats.fornecedor };

                const evolution = itemRef.weight - itemCurrent.weight;
                const isNew = !refItem && currentItem;
                const isDiscontinued = refItem && !currentItem;

                let presenceStatus = 'Ativo';
                let reason = '-';

                if (isNew) {
                    const existedBefore = stats.history.some(h => h.date < currentSnapshot.date && !h.missing);
                    if (existedBefore) {
                        presenceStatus = 'Reativado';
                        reason = 'Retornou ao giro';
                    } else {
                        presenceStatus = 'Novo';
                        reason = 'Entrou no giro';
                    }
                } else if (isDiscontinued) {
                    presenceStatus = 'Fora da Janela';
                    reason = 'Saiu da janela';
                } else {
                    presenceStatus = 'Ativo';
                    // Analisar motivo da evolução para itens que se mantiveram
                    if (evolution !== 0) {
                        const estAumentou = itemCurrent.estoque > itemRef.estoque;
                        const venCaiu = itemCurrent.venda < itemRef.venda;
                        if (estAumentou && venCaiu) reason = 'Estoque aumentou e Venda caiu';
                        else if (estAumentou) reason = 'Estoque aumentou';
                        else if (venCaiu) reason = 'Média de venda caiu';
                        else if (itemCurrent.estoque < itemRef.estoque && itemCurrent.venda > itemRef.venda) reason = 'Estoque caiu e Venda aumentou';
                        else if (itemCurrent.estoque < itemRef.estoque) reason = 'Estoque caiu';
                        else if (itemCurrent.venda > itemRef.venda) reason = 'Média de venda aumentou';
                    }
                }

                return { code: stats.code, itemRef, itemCurrent, baseItem, evolution, stats, isNew, isDiscontinued, presenceStatus, reason };
            })
            .filter(row => {
                const matchSearch = row.code.toLowerCase().includes(searchTerm) || row.stats.desc.toLowerCase().includes(searchTerm);
                if (!matchSearch) return false;

                if (filterVal === 'improved') return row.stats.histRecoveredAtCurrentIdx === true;
                if (filterVal === 'improved_rupture') return row.stats.histRecoveredAtCurrentIdx === true && row.baseItem.status === 'rupture';
                if (filterVal === 'improved_attention') return row.stats.histRecoveredAtCurrentIdx === true && row.baseItem.status === 'attention';
                if (filterVal === 'improved_suggest') return row.stats.histRecoveredAtCurrentIdx === true && row.baseItem.status === 'suggest';
                if (filterVal === 'worsened') return row.evolution < 0;
                if (filterVal === 'stable') return row.evolution === 0;
                if (filterVal === 'relapsed') return row.stats.isRecurrentAtCurrentIdx === true;
                if (filterVal === 'new_items') return row.isNew;
                if (filterVal === 'discontinued') return row.isDiscontinued;
                
                // Para 'Todos os Itens' (all), mostramos a lista completa de itens
                return true;
            });

        // Atualizar o título do painel para mostrar a quantidade de itens listados
        const tableEl = document.getElementById('evolution-table');
        const panelTitle = tableEl ? tableEl.closest('.main-card').querySelector('h3') : null;
        if (panelTitle) {
            panelTitle.innerHTML = `<span class="header-icon">📋</span> Detalhamento de Movimentação por Item (${rows.length} itens)`;
        }

        tableBody.innerHTML = rows.map(row => {
            const desc = String(row.stats.desc || 'S/D');
            const fornecedor = String(row.stats.fornecedor || 'N/D');
            const code = String(row.code || '0');
            
            let iconEvol = '';
            let colorEvol = 'color: #9ca3af;';
            if (row.evolution > 0) { iconEvol = '↑ ' + Math.abs(row.evolution); colorEvol = 'color: #34d399;'; }
            else if (row.evolution < 0) { iconEvol = '↓ ' + Math.abs(row.evolution); colorEvol = 'color: #fb7185;'; }
            else { iconEvol = '='; }

            let presenceBadge = '';
            if (row.presenceStatus === 'Novo') presenceBadge = '<span style="background: rgba(59, 130, 246, 0.1); color: #60a5fa; padding: 2px 6px; border-radius: 4px; font-size: 0.7rem; font-weight: bold;">✨ Novo</span>';
            else if (row.presenceStatus === 'Fora da Janela') presenceBadge = '<span style="background: rgba(107, 114, 128, 0.1); color: #9ca3af; padding: 2px 6px; border-radius: 4px; font-size: 0.7rem; font-weight: bold;">🗑️ Fora</span>';
            else if (row.presenceStatus === 'Reativado') presenceBadge = '<span style="background: rgba(168, 85, 247, 0.1); color: #c084fc; padding: 2px 6px; border-radius: 4px; font-size: 0.7rem; font-weight: bold;">♻️ Reativado</span>';
            else presenceBadge = '<span style="color: #94a3b8; font-size: 0.75rem;">Ativo</span>';

            const statusEvolText = row.evolution > 0 ? 'Melhorou' : (row.evolution < 0 ? 'Piorou' : 'Estável');
            const statusEvolClass = row.evolution > 0 ? 'ev-good' : (row.evolution < 0 ? 'ev-bad' : 'ev-stable');

            return `
            <tr>
                <td style="font-weight:600">${code}</td>
                <td style="font-size:0.8rem" title="${desc}">${desc.substring(0, 35)}${desc.length > 35 ? '...' : ''}</td>
                <td style="font-size:0.75rem; color:var(--text-muted)" title="${fornecedor}">${fornecedor.substring(0, 15)}${fornecedor.length > 15 ? '...' : ''}</td>
                <td style="text-align:center"><span class="badge badge-${row.itemRef.status}">${CRITICIDADE[row.itemRef.status].label}</span></td>
                <td style="text-align:center"><span class="badge badge-${row.itemCurrent.status}">${CRITICIDADE[row.itemCurrent.status].label}</span></td>
                <td style="text-align:center; font-weight:800; ${colorEvol}">${iconEvol}</td>
                <td style="text-align:center;">${presenceBadge}</td>
                <td style="font-size:0.75rem; color: #94a3b8;">${row.reason}</td>
                <td>
                    <span class="evolution-status-tag ${statusEvolClass}">
                        ${statusEvolText}
                    </span>
                </td>
                <td style="text-align:center;">
                    <button class="filter-btn btn-show-history" data-code="${code}" style="height:28px; padding:0 8px; font-size:0.7rem; border-color: rgba(99, 102, 241, 0.3);">
                        🔍 Histórico
                    </button>
                </td>
            </tr>
            `; 
        }).join('');
    }

    window.toggleSupplierFilter = (supplier) => {
        if (activeSupplier === supplier) {
            activeSupplier = 'all';
        } else {
            activeSupplier = supplier;
        }
        updateDashboard();
    };

    window.clearSupplierFilter = () => {
        activeSupplier = 'all';
        updateDashboard();
    };

    function updatePerformance() {
        const supplierPerfList = document.getElementById('supplier-performance-list');
        const cardHeader = supplierPerfList ? supplierPerfList.closest('.main-card').querySelector('h3') : null;
        
        if (cardHeader) {
            if (activeSupplier !== 'all') {
                cardHeader.innerHTML = `<span class="header-icon">🤝</span> Fornecedores <span class="badge" style="font-size: 0.65rem; background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.4); padding: 2px 6px; border-radius: 4px; margin-left: 8px; cursor: pointer;" onclick="window.clearSupplierFilter()">Filtro: ${activeSupplier.substring(0, 15)}${activeSupplier.length > 15 ? '...' : ''} ✕</span>`;
            } else {
                cardHeader.innerHTML = `<span class="header-icon">🤝</span> Fornecedores (Estabilidade)`;
            }
        }

        const suppliers = {};
        
        let refSnapshot = baseSnapshot;
        if (comparisonMode === 'diaria') {
            refSnapshot = currentTimelineIdx > 0 ? snapshotHistory[currentTimelineIdx - 1] : currentSnapshot;
        }

        const refItems = Array.from(refSnapshot.itemsMap.values());
        
        refItems.forEach(iRef => {
            const iCurr = currentSnapshot.itemsMap.get(iRef.code) || { weight: 0 };
            // A criticidade diminuiu significa melhoria/recuperação
            if (iRef.weight - iCurr.weight > 0) {
                suppliers[iRef.fornecedor] = (suppliers[iRef.fornecedor] || 0) + 1;
            }
        });

        // Fornecedores
        if (supplierPerfList) {
            const sortedSuppliers = Object.entries(suppliers).sort((a, b) => b[1] - a[1]);
            if (sortedSuppliers.length === 0) {
                supplierPerfList.innerHTML = `<div style="text-align:center; color:var(--text-muted); font-size:0.8rem; padding: 1rem 0;">Nenhuma recuperação</div>`;
            } else {
                supplierPerfList.innerHTML = sortedSuppliers.slice(0, 5).map(([name, count], idx) => {
                    const isSelected = name === activeSupplier;
                    const percent = (count / sortedSuppliers[0][1] * 100);
                    return `
                    <div class="perf-item ${isSelected ? 'selected' : ''}" 
                         onclick="window.toggleSupplierFilter('${name.replace(/'/g, "\\'")}')"
                         style="cursor: pointer; transition: all 0.2s; padding: 6px 8px; border-radius: 6px; border: 1px solid ${isSelected ? 'rgba(16, 185, 129, 0.4)' : 'transparent'}; background: ${isSelected ? 'rgba(16, 185, 129, 0.08)' : 'transparent'}; margin-bottom: 2px;">
                        <div class="perf-rank" style="${isSelected ? 'background: #10b981; color: white;' : ''}">${idx + 1}</div>
                        <div class="perf-info" style="flex: 1;">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                                <div class="perf-name" style="font-size: 0.8rem; font-weight: 600; color: ${isSelected ? '#34d399' : '#fff'};" title="${name}">${name}</div>
                                <div class="perf-sub" style="font-size: 0.7rem; font-weight: 700; color: ${isSelected ? '#34d399' : 'var(--text-muted)'};">${count} recup.</div>
                            </div>
                            <div class="perf-bar-wrapper" style="width: 100%; height: 6px; background: rgba(255, 255, 255, 0.05); border-radius: 3px; overflow: hidden;">
                                <div class="perf-bar" style="width: ${percent}%; height: 100%; background: linear-gradient(to right, #10b981, #34d399); border-radius: 3px;"></div>
                            </div>
                        </div>
                    </div>
                    `;
                }).join('');
            }
        }
    }

    function updateBuyerFilter() {
        // Removido a pedido do usuário
    }

    // Modal de Histórico do Item
    let lastOpenTime = 0;

    window.showItemHistory = (code) => {
        const stats = itemStatsMap.get(code);
        if (!stats) return alert("Histórico detalhado não localizado para o item " + code);

        document.getElementById('hist-item-code').textContent = `Produto: ${stats.code}`;
        document.getElementById('hist-item-desc').textContent = stats.desc;
        document.getElementById('hist-critical-snapshots').textContent = `${stats.criticalDays} snapshots em risco`;
        
        document.getElementById('hist-rupture-count').textContent = stats.ruptureCount;
        document.getElementById('hist-attention-count').textContent = stats.attentionCount;
        document.getElementById('hist-suggest-count').textContent = stats.suggestCount;
        
        document.getElementById('hist-recovered-count').textContent = stats.recuperacoesCount;
        document.getElementById('hist-worsened-count').textContent = stats.worsenedCount;

        const relapsedBadge = document.getElementById('hist-reincidente-badge');
        if (stats.reincidente) {
            relapsedBadge.style.display = 'block';
        } else {
            relapsedBadge.style.display = 'none';
        }

        const timelineList = document.getElementById('item-timeline-list');
        timelineList.innerHTML = stats.history.map(h => {
            const formattedDate = h.date.split('-').reverse().join('/');
            const statusLabel = CRITICIDADE[h.status].label;
            const statusColor = CRITICIDADE[h.status].color;
            
            return `
                <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.03); padding:8px 12px; border-radius:6px; border-left:4px solid ${statusColor};">
                    <span style="font-size:0.8rem; font-weight:500;">${formattedDate}</span>
                    <span class="badge badge-${h.status}" style="font-size:0.7rem; font-weight:700; background: ${statusColor}25; color: ${statusColor}; border: 1px solid ${statusColor}40;">${statusLabel}</span>
                </div>
            `;
        }).join('');

        const modal = document.getElementById('item-history-modal');
        modal.style.display = 'flex';
        setTimeout(() => {
            modal.classList.add('active');
        }, 10);
        lastOpenTime = Date.now(); // ⏱️ Registra o momento exato da abertura
    };

    const modal = document.getElementById('item-history-modal');
    const closeBtn = document.getElementById('close-item-history-btn');
    
    const closeModal = () => {
        if (Date.now() - lastOpenTime < 400) return; // 🛡️ Ignora cliques fantasmas em menos de 400ms
        modal.classList.remove('active');
        setTimeout(() => {
            if (!modal.classList.contains('active')) {
                modal.style.display = 'none';
            }
        }, 300); // Aguarda a transição de opacidade do CSS terminar
    };

    if (closeBtn) {
        closeBtn.onclick = closeModal;
    }
    
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            closeModal();
        }
    });

    // Listeners gerais
    folderInputs.forEach(input => {
        if (input) {
            input.onchange = (e) => handleFiles(e.target.files);
        }
    });

    timelineRange.oninput = (e) => {
        currentTimelineIdx = parseInt(e.target.value);
        updateDashboard();
    };

    tableSearch.oninput = renderTable;
    evolutionFilter.onchange = () => {
        renderTable();
        renderCharts();
    };

    tableBody.onclick = (e) => {
        const btn = e.target.closest('.btn-show-history');
        if (btn) {
            e.stopPropagation();
            const code = btn.dataset.code;
            window.showItemHistory(code);
        }
    };

    const cardFilters = {
        'card-recovery': 'improved',
        'card-rupture-recovered': 'improved_rupture',
        'card-attention-recovered': 'improved_attention',
        'card-suggest-recovered': 'improved_suggest',
        'card-relapsed': 'relapsed',
        'card-stable': 'stable',
        'card-worsened': 'worsened',
        'card-new-items': 'new_items',
        'card-discontinued': 'discontinued'
    };

    Object.entries(cardFilters).forEach(([cardId, filterVal]) => {
        const cardEl = document.getElementById(cardId);
        if (cardEl) {
            cardEl.addEventListener('click', () => {
                const isSelected = cardEl.classList.contains('selected');
                
                // Remover selected de todos os cards
                Object.keys(cardFilters).forEach(id => {
                    const el = document.getElementById(id);
                    if (el) el.classList.remove('selected');
                });

                if (isSelected) {
                    // Desmarcar: voltar para "Todos os Itens"
                    evolutionFilter.value = 'all';
                } else {
                    // Marcar: aplicar o filtro correspondente
                    cardEl.classList.add('selected');
                    evolutionFilter.value = filterVal;
                }

                // Disparar atualização da tabela e do gráfico
                renderTable();
                renderCharts();
            });
        }
    });

    // Sincronizar o destaque visual do card se o usuário mudar o filtro através do select
    evolutionFilter.addEventListener('change', () => {
        const currentVal = evolutionFilter.value;
        Object.entries(cardFilters).forEach(([cardId, filterVal]) => {
            const el = document.getElementById(cardId);
            if (el) {
                if (filterVal === currentVal) {
                    el.classList.add('selected');
                } else {
                    el.classList.remove('selected');
                }
            }
        });
        renderCharts();
    });
});
