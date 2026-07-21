document.addEventListener('DOMContentLoaded', () => {
    console.log("Gestor de Estoque v2.2 - Sincronização Dinâmica Ativa");
    const fileUpload = document.getElementById('file-upload');
    const dropZone = document.getElementById('drop-zone');
    const statsSection = document.getElementById('stats-section');
    const tableContainer = document.getElementById('table-container');
    const tableBody = document.getElementById('table-body');
    const tableSearch = document.getElementById('table-search');
    const smartSearch = document.getElementById('smart-desc-search');
    const buyerBtns = document.querySelectorAll('.buyer-btn');
    const buyerUpload = document.getElementById('buyer-upload');
    const filterBtns = document.querySelectorAll('.filter-btn');
    const clearFiltersBtn = document.getElementById('clear-filters');
    const chartSection = document.getElementById('chart-section');
    const loadingOverlay = document.getElementById('loading-overlay');
    const folderUpload = document.getElementById('folder-upload');
    const supplierModal = document.getElementById('supplier-modal');
    const openSupplierModalBtn = document.getElementById('open-supplier-modal');
    const closeSupplierModalBtn = document.getElementById('close-supplier-modal');
    const applySupplierFilterBtn = document.getElementById('apply-supplier-filter');
    const supplierChecklist = document.getElementById('supplier-checklist');
    const modalSupplierSearch = document.getElementById('modal-supplier-search');
    const btnSelectAllSuppliers = document.getElementById('btn-select-all-suppliers');
    const btnClearAllSuppliers = document.getElementById('btn-clear-all-suppliers');

    // --- Loading Control ---
    function showLoading() {
        if (loadingOverlay) loadingOverlay.style.display = 'flex';
        document.body.classList.add('loading-active');
    }

    function hideLoading() {
        if (loadingOverlay) loadingOverlay.style.display = 'none';
        document.body.classList.remove('loading-active');
    }


    // Modal Logic
    const rulesModal = document.getElementById('rules-modal');
    const openRulesBtn = document.getElementById('open-rules-btn');
    const closeRulesBtn = document.getElementById('close-rules-btn');

    if (openRulesBtn && closeRulesBtn && rulesModal) {
        openRulesBtn.addEventListener('click', () => {
            rulesModal.classList.add('active');
        });
        closeRulesBtn.addEventListener('click', () => {
            rulesModal.classList.remove('active');
        });
        rulesModal.addEventListener('click', (e) => {
            if (e.target === rulesModal) rulesModal.classList.remove('active');
        });
    }

    if (folderUpload) folderUpload.addEventListener('change', handleFolderUpload);

    // Register ChartJS Plugin
    Chart.register(ChartDataLabels);

    let currentData = [];
    let filteredData = [];
    let activeFilters = ['all'];
    let activeBuyer = 'all';
    let activeRecFilter = null; // Filter by recurrence bracket (only for Ruptura)
    let sortRecorrenciaDir = 'none';
    let sortVendasDir = 'none';
    let sortDiasEstoqueDir = 'none';
    let isRecurrenceActive = false;
    let myChart = null;
    let supplierChart = null;
    let groupChart = null;
    let selectedSuppliers = [];
    let selectedGroups = [];
    let fixedSupplierLabels = [];
    let fixedGroupLabels = [];

    let buyerMap = JSON.parse(localStorage.getItem('buyerMap') || '{}');
    console.log(`Mapeamento de compradores carregado: ${Object.keys(buyerMap).length} códigos.`);

    let leadtimeSalesHistory = {};

    /**
     * Busca histórico de vendas oficial do Leadtime via Bridge
     */
    async function fetchSalesHistory() {
        try {
            console.log("Tentando buscar histórico do Leadtime via Bridge...");
            const response = await fetch(`http://localhost:3000/sales-history?t=${Date.now()}`);
            const result = await response.json();
            if (result.success) {
                leadtimeSalesHistory = result.data;
                console.log(`✅ Sucesso! ${Object.keys(leadtimeSalesHistory).length} itens carregados do Leadtime.`);
                if (currentData.length > 0) {
                    console.log("Re-processando dados com novo histórico do Leadtime...");
                    // Se já houver dados, re-processamos com o novo histórico
                    // Como processData precisa do JSON bruto, e não o temos salvo, 
                    // o ideal seria que o usuário fizesse o upload novamente ou tivéssemos o JSON original.
                    // Para simplificar, informamos que o Leadtime está pronto.
                }
            }
        } catch (e) {
            console.warn("⚠️ Bridge indisponível. Use o botão 'Sync OneDrive' para carregar o histórico manualmente.");
        }
    }

    /**
     * Processa a pasta do OneDrive selecionada pelo usuário
     */
    async function handleFolderUpload(e) {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        showLoading();
        try {
            // Procurar especificamente pelo arquivo de Leadtime
            const leadtimeFile = files.find(f => f.name.toLowerCase().includes('base leadtime') && f.name.endsWith('.xlsx'));
            
            if (leadtimeFile) {
                console.log(`Arquivo Leadtime encontrado: ${leadtimeFile.name}`);
                const data = await readExcel(leadtimeFile);
                processLeadtimeData(data);
                alert(`✅ Histórico do Leadtime sincronizado (${Object.keys(leadtimeSalesHistory).length} itens).`);
                
                // Se já tivermos dados na tabela, precisamos atualizar as médias e cores
                if (currentData.length > 0) {
                    // Nota: Idealmente salvaríamos o JSON bruto original para re-processar.
                    // Como solução paliativa, informamos ao usuário para carregar a planilha de giro novamente
                    // ou tentamos atualizar os objetos existentes se o código bater.
                    updateExistingDataWithLeadtime();
                }
            } else {
                alert("Arquivo 'Base Leadtime.xlsx' não encontrado na pasta selecionada.");
            }
        } catch (err) {
            console.error("Erro ao processar pasta OneDrive:", err);
            alert("Erro ao processar arquivos da pasta.");
        } finally {
            hideLoading();
            e.target.value = '';
        }
    }

    function readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                resolve(json);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function processLeadtimeData(rows) {
        if (!rows || rows.length === 0) return;

        // Identificar a linha de cabeçalho e as colunas de meses (mesma lógica do bridge.js)
        let headerRow = rows[0];
        for (let i = 0; i < Math.min(rows.length, 10); i++) {
            if (rows[i].some(cell => {
                if (cell instanceof Date) return true;
                return typeof cell === 'string' && /^\d{2}\/\d{4}$/.test(cell);
            })) {
                headerRow = rows[i];
                break;
            }
        }

        const monthCols = headerRow
            .map(k => {
                if (k instanceof Date) {
                    const m = (k.getMonth() + 1).toString().padStart(2, '0');
                    const y = k.getFullYear();
                    return `${m}/${y}`;
                }
                return String(k || '');
            })
            .filter(k => /^\d{2}\/\d{4}$/.test(k))
            .filter(m => {
                const [mon, yr] = m.split('/').map(Number);
                return (yr < 2026) || (yr === 2026 && mon <= 5);
            })
            .sort((a, b) => {
                const [mA, yA] = a.split('/').map(Number);
                const [mB, yB] = b.split('/').map(Number);
                return (yA * 12 + mA) - (yB * 12 + mB);
            });

        const uniqueMonthCols = [...new Set(monthCols)];
        const monthMap = [];
        headerRow.forEach((val, idx) => {
            let label = "";
            if (val instanceof Date) {
                const m = (val.getMonth() + 1).toString().padStart(2, '0');
                const y = val.getFullYear();
                label = `${m}/${y}`;
            } else {
                label = String(val || '');
            }
            if (uniqueMonthCols.includes(label)) {
                monthMap.push({ index: idx, label: label });
            }
        });

        monthMap.sort((a, b) => {
            const [mA, yA] = a.label.split('/').map(Number);
            const [mB, yB] = b.label.split('/').map(Number);
            return (yA * 12 + mA) - (yB * 12 + mB);
        });

        const salesMap = {};
        const dataRows = rows.slice(rows.indexOf(headerRow) + 1);
        const prodIdx = headerRow.findIndex(h => h && h.toString().toLowerCase().includes('produto'));
        
        dataRows.forEach(row => {
            const code = String(row[prodIdx] || '').trim();
            if (!code || code === 'undefined') return;

            const history = monthMap.map(m => parseFloat(row[m.index]) || 0);
            salesMap[code] = {
                history: history,
                labels: monthMap.map(m => m.label)
            };
        });

        leadtimeSalesHistory = salesMap;
    }

    function updateExistingDataWithLeadtime() {
        // Atualiza os dados já carregados na memória com o novo histórico do Leadtime
        currentData.forEach(item => {
            const code = item.produto;
            if (code !== 'N/A' && leadtimeSalesHistory[code]) {
                const lt = leadtimeSalesHistory[code];
                item.historico = lt.history;
                item.monthLabels = lt.labels;
                item.vendas = lt.history.reduce((a, b) => a + b, 0);
                
                const activeMonths = lt.history.filter(v => v > 0).length;
                const finalRecorrencia = lt.history.length > 0 ? (activeMonths / lt.history.length) * 100 : 0;
                const finalMedVenda = activeMonths > 0 ? (item.vendas / activeMonths) : 0;
                
                item.medVenda = finalMedVenda.toFixed(3);
                item.recorrencia = finalRecorrencia.toFixed(0);
                
                // Recalcular situação e dias de estoque
                const DIAS_UTEIS_MES = 22;
                item.diasEstoque = (finalMedVenda > 0) ? Math.round((item.estoque / finalMedVenda) * DIAS_UTEIS_MES) : null;
                
                item.situacao = 'seguro';
                item.emRisco = false;
                item.emAtencao = false;
                item.emSugestao = false;
                
                if (finalRecorrencia > 17) {
                    const totalDisponivel = item.estoque + item.encomendas;
                    if (finalMedVenda > totalDisponivel) {
                        item.situacao = 'ruptura'; item.emRisco = true;
                    } else if ((finalMedVenda * 2) > totalDisponivel) {
                        item.situacao = 'atencao'; item.emAtencao = true;
                    } else if ((finalMedVenda * 3) > totalDisponivel) {
                        item.situacao = 'sugestao'; item.emSugestao = true;
                    }
                }
            }
        });
        
        filteredData = [...currentData];
        renderTable(filteredData);
    }

    fetchSalesHistory();

    // --- Core Logic ---

    /**
     * Normalizes Unit of Measure to 'Pç'
     */
    function normalizeUnit(un) {
        if (!un) return 'Pç';
        const u = un.toString().toUpperCase().trim();
        if (u === 'PC' || u === 'UN' || u === 'UND' || u === 'PÇ' || u === 'PEÇA') return 'Pç';
        return u;
    }

    /**
     * Tenta encontrar o valor em uma linha usando múltiplos nomes de coluna possíveis (Case-Insensitive)
     */
    function getValue(row, keys) {
        const rowKeys = Object.keys(row);
        const cleanStr = (s) => s.toString().toLowerCase().trim()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/[.\-_/\\()]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

        for (let key of keys) {
            const target = cleanStr(key);
            
            // 1. Exact match in raw keys
            if (row[key] !== undefined) return { value: row[key], col: key };
            
            // 2. Exact match in normalized keys
            const normKey = rowKeys.find(rk => cleanStr(rk) === target);
            if (normKey) return { value: row[normKey], col: normKey };
        }

        // 3. Fallback: word match or partial match
        for (let key of keys) {
            const target = cleanStr(key);
            if (target.length < 3) continue;
            
            const fuzzyKey = rowKeys.find(rk => {
                const cRK = cleanStr(rk);
                const words = cRK.split(' ');
                
                // Matches if any word in the column name starts with our target or vice versa
                return words.some(word => {
                    if (word === target) return true;
                    if (word.length >= 4 && target.startsWith(word)) return true;
                    if (target.length >= 4 && word.startsWith(target)) return true;
                    return false;
                });
            });
            if (fuzzyKey) return { value: row[fuzzyKey], col: fuzzyKey };
        }
        
        return { value: undefined, col: 'N/A' };
    }

    /**
     * Converte valor para número de forma segura, tratando formatos brasileiros (vírgula/ponto)
     */
    function parseNum(val) {
        if (val === undefined || val === null || val === '') return 0;
        if (typeof val === 'number') return val;
        let str = val.toString().replace('R$', '').replace(/\s/g, '').trim();
        
        // Handle BR/INT formatting automatically
        if (str.startsWith('.') || str.startsWith(',')) str = '0' + str;

        const hasComma = str.includes(',');
        const hasDot = str.includes('.');

        if (hasComma && hasDot) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
                str = str.replace(/\./g, '').replace(',', '.'); // BR
            } else {
                str = str.replace(/,/g, ''); // INT
            }
        } else if (hasComma) {
            str = str.replace(',', '.');
        }
        
        const num = parseFloat(str);
        // For inventory/sales, we generally treat negative/NaN as 0 to avoid visual artifacts
        return isNaN(num) ? 0 : num;
    }

    /**
     * Calculates Recurrence % (Months with sales > 0)
     */
    function calculateRecurrence(row, monthCols) {
        let count = 0;
        if (monthCols && monthCols.length > 0) {
            monthCols.forEach(col => {
                const val = parseNum(row[col]);
                if (val > 0) count++;
            });
        }
        const percent = monthCols.length > 0 ? (count / monthCols.length) * 100 : 0;
        return { percent, count };
    }

    /**
     * Processes raw JSON from Excel
     */
    function processData(rawData) {
        if (!rawData || rawData.length === 0) return [];

        // IMPROVED: Gather all unique keys across all rows to avoid missing columns
        // if the first row is incomplete.
        const allKeysSet = new Set();
        rawData.forEach(row => {
            Object.keys(row).forEach(k => allKeysSet.add(k));
        });
        const allKeys = Array.from(allKeysSet);

        // Identify Month Columns (format MM/YYYY) - Dynamic Sorting
        // Filtrar apenas o que realmente parece data MM/YYYY e ignorar lixo ou meses futuros
        const monthCols = allKeys.filter(k => {
            if (!/^\d{2}\/\d{4}$/.test(k)) return false;
            const [mon, yr] = k.split('/').map(Number);
            // Excluir estritamente qualquer mês após 05/2026 conforme solicitado
            return (yr < 2026) || (yr === 2026 && mon <= 5);
        }).sort((a, b) => {
            const [mA, yA] = a.split('/').map(Number);
            const [mB, yB] = b.split('/').map(Number);
            return (yA * 12 + mA) - (yB * 12 + mB);
        });
        
        console.log(`Dados Brutos: ${rawData.length} linhas | Colunas Detectadas: ${allKeys.length}`);
        console.log('Meses identificados:', monthCols);

        const processed = [];
        
        rawData.forEach((row, index) => {
            const prodObj = getValue(row, ['Produto', 'Código', 'Item', 'Cód.', 'ID']);
            const descObj = getValue(row, ['Descrição longa do produto', 'Descrição', 'Desc', 'Nome', 'Produto Descrição', 'Texto']);
            
            // VALIDATION: Skip row if it has no product ID AND no description (likely empty or junk row)
            if (!prodObj.value && !descObj.value) {
                return; 
            }

            const produtoRaw = prodObj.value ? prodObj.value.toString().trim() : 'N/A';
            const produto = produtoRaw;
            const descricao = descObj.value ? descObj.value.toString().trim() : 'Sem descrição';

            // Lookup Buyer
            let comprador = 'N/D';
            if (produto !== 'N/A') {
                // Try exact match, then match without leading zeros/dots if numeric
                const cleanCode = produto.replace(/^0+/, '').replace(/[.]/g, '');
                comprador = buyerMap[produto] || buyerMap[cleanCode] || 'N/D';
            }

            const unObj = getValue(row, ['UN', 'Unidade', 'Medida', 'U', 'Med']);
            const grupoObj = getValue(row, ['Grupo', 'Categoria', 'Família', 'Linha', 'Grupo de Produto', 'Cód. Grupo', 'Subgrupo']);
            const fornecObj = getValue(row, ['Razão social do fornecedor', 'Fornecedor', 'Fornec', 'Fabricante', 'Último Fornecedor', 'Fornecedor Principal', 'Nome Fornecedor']);
            const vendasObj = getValue(row, ['Vendas', 'Qtd. Vendida', 'Venda Total', 'Total Vendas', 'Venda', 'Saídas', 'Giro']);
            const estoqueObj = getValue(row, ['Estoque', 'Saldo', 'Qtd. Estoque', 'Estoque Total', 'Saldo Atual', 'Saldo Disponível', 'Disp.', 'Qtd. Disponível', 'Estoque Atual']);
            const encomObj = getValue(row, [
                'Encomendas', 'Qtd. Encomenda', 'Saldo Pedido Compra', 'Saldo Ped. Compra', 'Pedido Compra', 
                'Qtd. em Pedido Compra', 'Qtd. no Pedido Compra', 'Saldo a Receber', 'A Receber', 'Pedidos', 
                'Qtd. Pedida', 'Saldo Pedido', 'Compras', 'Qtd em Pedido', 'Qtd. Ped.', 'Saldo Ped.', 
                'Pendência', 'Qtd. no Pedido', 'Encomenda', 'Pedido', 'Qtd Ped Compra', 'A Receber Total',
                'A Entregar', 'Saldo a Entregar', 'Qtd. Pendente', 'Pendente', 'Saldo O.C.', 'Ord. Compra'
            ]);
            const custoObj = getValue(row, ['Preço reposição', 'Custo aquisição', 'Custo Unitário', 'Custo', 'Preço Custo', 'Vlr. Custo', 'Custo Médio', 'Unitário']);

            let vendas = parseNum(vendasObj.value);
            
            // INTERPRETAÇÃO DA FÓRMULA EXCEL (Base Leadtime): 
            // Recorrência = CONT.SES(AB2:AG2;">0")/6
            const recData = calculateRecurrence(row, monthCols);
            
            // FALLBACK: If "Vendas" column is 0 or missing, sum the month columns
            if (vendas === 0 && monthCols.length > 0) {
                monthCols.forEach(col => {
                    const v = parseNum(row[col]);
                    if (v > 0) vendas += v;
                });
            }

            const medVendaFromHistory = (recData.count > 0 ? (vendas / recData.count) : 0);
            const recorrencia = recData.percent;
            
            // --- OVERRIDE WITH LEADTIME DATA IF AVAILABLE ---
            let finalHistorico = monthCols.map(col => Math.max(0, parseNum(row[col])));
            let finalMonthLabels = monthCols;
            let finalVendas = vendas;
            let finalRecorrencia = recorrencia;
            let finalMedVenda = medVendaFromHistory;

            if (produto !== 'N/A' && leadtimeSalesHistory[produto]) {
                const lt = leadtimeSalesHistory[produto];
                finalHistorico = lt.history;
                finalMonthLabels = lt.labels;
                finalVendas = lt.history.reduce((a, b) => a + b, 0);
                
                const activeMonths = lt.history.filter(v => v > 0).length;
                finalRecorrencia = lt.history.length > 0 ? (activeMonths / lt.history.length) * 100 : 0;
                finalMedVenda = activeMonths > 0 ? (finalVendas / activeMonths) : 0;
                
                // console.log(`   - Usando histórico Leadtime para ${produto}`);
            }

            const medVenda = finalMedVenda;
            
            const estoque = parseNum(estoqueObj.value);
            const encomendas = parseNum(encomObj.value);
            const custo = parseNum(custoObj.value);
            
            const currentRecorrencia = finalRecorrencia;

            // Dias úteis de estoque = (estoque / média mensal de venda) × 22
            // Brasil: ~22 dias úteis por mês (meses com 5 semanas = 21-23, média 22)
            const DIAS_UTEIS_MES = 22;
            const diasEstoque = (medVenda > 0) ? Math.round((estoque / medVenda) * DIAS_UTEIS_MES) : null;
            
            const isFromLeadtime = (produto !== 'N/A' && leadtimeSalesHistory[produto]);
            const mappingInfo = `Linha: ${index + 2} | Cód: "${prodObj.col}" | Estoque: "${estoqueObj.col}" | Encomendas: "${encomObj.col}"${isFromLeadtime ? ' | Histórico: Leadtime 📊' : ''}`;
            
            // Logic for risk classification (Refactored for maximum clarity)
            let emRisco = false;
            let emAtencao = false;
            let emSugestao = false;
            let situacao = 'seguro';

            // INTERPRETAÇÃO DA FÓRMULA EXCEL (Base Leadtime): 
            // Comprar = SE(E2>F2+G2;"Comprar";"Não comprar")
            // Onde E2 = Méd.Venda, F2 = Estoque, G2 = Encomendas
            
            // Only evaluate for risks if recurrence is > 33%
            if (currentRecorrencia > 17) {
                const totalDisponivel = estoque + encomendas;
                
                if (medVenda > totalDisponivel) {
                    emRisco = true;
                    situacao = 'ruptura';
                } else if ((medVenda * 2) > totalDisponivel) {
                    emAtencao = true;
                    situacao = 'atencao';
                } else if ((medVenda * 3) > totalDisponivel) {
                    emSugestao = true;
                    situacao = 'sugestao';
                }
            }

            let tendencia = 'stable';
            if (finalHistorico.length >= 2) {
                const lastVal = finalHistorico[finalHistorico.length - 1];
                const prevVal = finalHistorico[finalHistorico.length - 2];
                if (lastVal > prevVal) tendencia = 'up';
                else if (lastVal < prevVal) tendencia = 'down';
            }

            processed.push({
                produto,
                descricao,
                comprador,
                un: normalizeUnit(unObj.value),
                grupo: grupoObj.value || 'Geral',
                fornecedor: fornecObj.value || 'N/D',
                estoque,
                encomendas,
                vendas: finalVendas,
                medVenda: medVenda.toFixed(3),
                diasEstoque,
                tendencia,
                custoRaw: custo,
                custo: custo.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
                recorrencia: currentRecorrencia.toFixed(0),
                historico: finalHistorico,
                monthLabels: finalMonthLabels,
                situacao,
                emRisco,
                emAtencao,
                emSugestao,
                temEncomenda: encomendas > 0,
                mappingInfo
            });
        });

        // --- NOVO: Mapeamento de Itens Similares e Nacional (ZB/ZA -> POL/OX e " I" -> Nacional) ---
        const descMap = {};
        const fullDescMap = {}; // Mapeamento por descrição completa para encontrar o nacional
        processed.forEach(item => {
            const upperDesc = item.descricao.toUpperCase();
            if (!fullDescMap[upperDesc]) fullDescMap[upperDesc] = [];
            fullDescMap[upperDesc].push(item);

            // Remove sufixos comuns para encontrar a base da descrição
            const base = item.descricao.replace(/\s(ZB|ZA|OX|POL|POLIDO|ZINCADO\sBRANCO|ZINCADO\sAMARELO)(\sI)?$/i, '').trim().toUpperCase();
            item.baseDesc = base;
            if (!descMap[base]) descMap[base] = [];
            descMap[base].push(item);
        });

        processed.forEach(item => {
            const upperDesc = item.descricao.toUpperCase();
            
            // 1. Lógica para Equivalente Nacional (Se termina em " I")
            if (upperDesc.endsWith(' I')) {
                const nationalDesc = upperDesc.substring(0, upperDesc.length - 2).trim();
                if (fullDescMap[nationalDesc]) {
                    item.nationalEquivalent = fullDescMap[nationalDesc];
                }
            }

            // 2. Lógica para Opção Econômica (ZB/ZA -> POL/OX)
            const isZinc = /\s(ZB|ZA|ZINCADO\sBRANCO|ZINCADO\sAMARELO)(\sI)?$/i.test(upperDesc);
            
            if (isZinc) {
                // Filtra itens que têm a mesma base mas são POL ou OX
                item.relatedItems = descMap[item.baseDesc].filter(other => {
                    if (other.produto === item.produto) return false;
                    const otherUpper = other.descricao.toUpperCase();
                    // Deve ser POL ou OX e ter o mesmo status de importação (ou ambos serem nacionais/importados)
                    // Para simplificar, focamos em ser POL/OX
                    return /\s(POL|OX|POLIDO)$/i.test(otherUpper);
                });
            }
        });

        console.log(`Itens Processados: ${processed.length}`);
        return processed;
    }

    /**
     * Renders processed data to the table
     */
    function renderTable(data) {
        tableBody.innerHTML = '';
        data.forEach(item => {
            const mainTr = document.createElement('tr');
            mainTr.className = 'main-row';
            
            let badgeClass = 'badge-ok';
            let badgeText = '✅ SEGURO';
            if (item.situacao === 'ruptura') {
                badgeClass = 'badge-buy';
                badgeText = '🚨 RUPTURA';
            } else if (item.situacao === 'atencao') {
                badgeClass = 'badge-warn';
                badgeText = '⚡ ATENÇÃO';
            } else if (item.situacao === 'sugestao') {
                badgeClass = 'badge-caution';
                badgeText = '⚠️ COMPRAR';
            }
            
            // Highlight if there's a pending order
            if (item.temEncomenda) {
                badgeText += ' 📦';
            }
            
            mainTr.innerHTML = `
                <td>
                    <span class="cell-label">Produto</span>
                    <div class="cell-value" style="font-weight: 600;">${item.produto}</div>
                </td>
                <td>
                    <span class="cell-label">Descrição</span>
                    <div class="cell-value td-desc">${item.descricao}</div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">UN</span>
                    <div class="cell-value">${item.un}</div>
                </td>

                <td>
                    <span class="cell-label">Fornecedor</span>
                    <div class="cell-value td-supplier" title="${item.fornecedor}">${item.fornecedor}</div>
                </td>
                <td>
                    <span class="cell-label">Comprador</span>
                    <div class="cell-value">${item.comprador}</div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Estoque</span>
                    <div class="cell-value" style="display: flex; flex-direction: column; align-items: center;">
                        <span style="font-weight: 600;">${item.estoque}</span>
                        ${item.temEncomenda ? `<span style="font-size: 0.75rem; color: var(--info); font-weight: 700;" title="Saldo em Pedido">+ ${item.encomendas} ped.</span>` : ''}
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Dias Est.</span>
                    <div class="cell-value">
                        ${renderDiasEstoque(item.diasEstoque)}
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Vendas (6m)</span>
                    <div class="cell-value" style="display: flex; align-items: center; justify-content: center;">
                        <span>${item.vendas}</span>
                        ${item.tendencia === 'up' ? '<span class="trend-indicator trend-up" title="Vendas subindo em relação ao mês anterior">▲</span>' : 
                          item.tendencia === 'down' ? '<span class="trend-indicator trend-down" title="Vendas caindo em relação ao mês anterior">▼</span>' : ''}
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Recorrência</span>
                    <div class="cell-value" style="display: flex; flex-direction: column; align-items: center; gap: 4px;">
                        <div class="rec-bar-container" style="width: 60px;">
                            <div class="rec-bar" style="width: ${item.recorrencia}%"></div>
                        </div>
                        <span style="font-size: 0.75rem; font-weight: 600;">${item.recorrencia}%</span>
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Estado</span>
                    <div class="cell-value">
                        <span class="badge ${badgeClass}">${badgeText}</span>
                    </div>
                </td>
            `;
            
            const detailTr = document.createElement('tr');
            detailTr.className = 'detail-row hidden';
            detailTr.innerHTML = `
                <td colspan="10">
                    <div class="row-details">
                        <div style="display: flex; flex-direction: column; gap: 1.5rem; flex: 1 1 600px;">
                            <div style="display: flex; flex-wrap: wrap; gap: 2rem;">
                                <div class="detail-item">
                                    <span class="detail-label">Encomendas</span>
                                    <span class="detail-value">${item.encomendas}</span>
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Média Venda</span>
                                    <span class="detail-value">
                                        ${item.medVenda}
                                        ${item.tendencia === 'up' ? '<span class="trend-indicator trend-up" style="margin-left:5px;">▲</span>' : 
                                          item.tendencia === 'down' ? '<span class="trend-indicator trend-down" style="margin-left:5px;">▼</span>' : ''}
                                    </span>
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Cobertura de Estoque</span>
                                    <span class="detail-value">
                                        ${renderDiasEstoque(item.diasEstoque)}
                                    </span>
                                </div>
                                <div class="detail-item">
                                    <span class="detail-label">Custo Unit.</span>
                                    <span class="detail-value">${item.custo}</span>
                                </div>
                            </div>

                            ${item.relatedItems && item.relatedItems.length > 0 ? `
                            <div class="detail-item economic-option" style="margin-top: 0;">
                                <div class="economic-header">
                                    <span class="economic-icon">💡</span>
                                    <div class="economic-titles">
                                        <span class="economic-badge">Alternativa</span>
                                        <span class="economic-subtitle">(Versão Polida/OX)</span>
                                    </div>
                                </div>
                                <div class="economic-list">
                                    ${item.relatedItems.map(rel => `
                                        <div class="economic-row">
                                            <div class="economic-info">
                                                <span class="economic-name">${rel.descricao}</span>
                                                <span class="economic-code">Cód: ${rel.produto}</span>
                                            </div>
                                            <div class="economic-cost">
                                                <span class="cost-value">${rel.custo}</span>
                                                <span class="cost-label">Custo Unit.</span>
                                            </div>
                                            <div class="economic-sales" style="text-align: right; display: flex; flex-direction: column; margin-right: 12px; min-width: 80px;">
                                                <span class="sales-value" style="font-size: 0.85rem; font-weight: 700; color: var(--text-main);">${rel.vendas}</span>
                                                <span class="sales-label" style="font-size: 0.6rem; color: var(--text-muted); text-transform: uppercase;">Vendas (6m)</span>
                                            </div>
                                            <div class="economic-stock">
                                                <span class="stock-value">${rel.estoque} <small>${rel.un}</small></span>
                                                <span class="stock-label">Em Estoque</span>
                                            </div>
                                            <div class="economic-coverage">
                                                ${renderDiasEstoque(rel.diasEstoque)}
                                            </div>
                                        </div>
                                    `).join('')}
                                </div>
                            </div>
                            ` : ''}

                            ${item.nationalEquivalent && item.nationalEquivalent.length > 0 ? `
                            <div class="detail-item national-option" style="margin-top: 0;">
                                <div class="economic-header">
                                    <span class="economic-icon">🇧🇷</span>
                                    <div class="economic-titles">
                                        <span class="national-badge">Equivalente Nacional</span>
                                        <span class="economic-subtitle">(Sugestão sem "I")</span>
                                    </div>
                                </div>
                                <div class="economic-list">
                                    ${item.nationalEquivalent.map(nat => `
                                        <div class="economic-row national">
                                            <div class="economic-info">
                                                <span class="economic-name">${nat.descricao}</span>
                                                <span class="economic-code">Cód: ${nat.produto}</span>
                                            </div>
                                            <div class="economic-cost">
                                                <span class="cost-value">${nat.custo}</span>
                                                <span class="cost-label">Custo Unit.</span>
                                            </div>
                                            <div class="economic-sales" style="text-align: right; display: flex; flex-direction: column; margin-right: 12px; min-width: 80px;">
                                                <span class="sales-value" style="font-size: 0.85rem; font-weight: 700; color: var(--text-main);">${nat.vendas}</span>
                                                <span class="sales-label" style="font-size: 0.6rem; color: var(--text-muted); text-transform: uppercase;">Vendas (6m)</span>
                                            </div>
                                            <div class="economic-stock">
                                                <span class="stock-value">${nat.estoque} <small>${nat.un}</small></span>
                                                <span class="stock-label">Em Estoque</span>
                                            </div>
                                            <div class="economic-coverage">
                                                ${renderDiasEstoque(nat.diasEstoque)}
                                            </div>
                                        </div>
                                    `).join('')}
                                </div>
                            </div>
                            ` : ''}
                        </div>

                        <div class="detail-item sparkline-detail">
                            <span class="detail-label">Histórico Vendas (Mensal)</span>
                            <div class="sparkline-container">
                                ${generateSparkline(item.historico, item.monthLabels)}
                            </div>
                        </div>

                        <div class="detail-item" style="border-top:1px solid var(--border-color); padding-top:10px; width:100%;">
                            <span class="detail-label">Mapeamento de Dados (Debug)</span>
                            <span class="detail-value" style="font-size:0.75rem; color:var(--text-muted); font-weight:normal;">
                                ${item.mappingInfo}
                            </span>
                        </div>
                    </div>
                </td>
            `;

            mainTr.addEventListener('click', () => {
                const isHidden = detailTr.classList.contains('hidden');
                // Optional: Close other expanded rows
                // document.querySelectorAll('.detail-row').forEach(dr => dr.classList.add('hidden'));
                // document.querySelectorAll('.main-row').forEach(mr => mr.classList.remove('expanded'));

                if (isHidden) {
                    detailTr.classList.remove('hidden');
                    mainTr.classList.add('expanded');
                } else {
                    detailTr.classList.add('hidden');
                    mainTr.classList.remove('expanded');
                }
            });

            tableBody.appendChild(mainTr);
            tableBody.appendChild(detailTr);
        });
    }

    /**
     * Updates the dashboard cards and pie chart based on the provided data
     */
    function updateDashboardUI(data) {
        const formatValue = (val) => val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        
        const buyItems = data.filter(i => i.situacao === 'ruptura');
        const attentionItems = data.filter(i => i.situacao === 'atencao');
        const suggestItems = data.filter(i => i.situacao === 'sugestao');
        const safeItems = data.filter(i => i.situacao === 'seguro');
        const orderItems = data.filter(i => i.temEncomenda);

        const sumValAlerts = (items, m) => items.reduce((acc, i) => acc + ((parseFloat(i.medVenda) || 0) * 3 * m * (parseFloat(i.custoRaw) || 0)), 0);
        const sumValStock = (items) => items.reduce((acc, i) => acc + (i.estoque * (parseFloat(i.custoRaw) || 0)), 0);
        const sumOrderVal = (items) => items.reduce((acc, i) => acc + (i.encomendas * (parseFloat(i.custoRaw) || 0)), 0);

        document.getElementById('total-items').textContent = data.length;
        const totalValueEl = document.getElementById('total-value');
        if (totalValueEl) totalValueEl.textContent = formatValue(sumValStock(data));

        document.getElementById('to-buy-count').textContent = buyItems.length;
        const buyValueEl = document.getElementById('buy-value');
        if (buyValueEl) buyValueEl.textContent = formatValue(sumValAlerts(buyItems, 1));

        document.getElementById('attention-count').textContent = attentionItems.length;
        const attentionValueEl = document.getElementById('attention-value');
        if (attentionValueEl) attentionValueEl.textContent = formatValue(sumValAlerts(attentionItems, 2));

        document.getElementById('suggestion-count').textContent = suggestItems.length;
        const suggestionValueEl = document.getElementById('suggestion-value');
        if (suggestionValueEl) suggestionValueEl.textContent = formatValue(sumValAlerts(suggestItems, 3));

        const safeCountEl = document.getElementById('safe-count');
        if (safeCountEl) safeCountEl.textContent = safeItems.length;
        const safeValueEl = document.getElementById('safe-value');
        if (safeValueEl) safeValueEl.textContent = formatValue(sumValStock(safeItems));

        const ordersCountEl = document.getElementById('orders-count');
        if (ordersCountEl) ordersCountEl.textContent = orderItems.length;
        const ordersValueEl = document.getElementById('orders-value');
        if (ordersValueEl) ordersValueEl.textContent = formatValue(sumOrderVal(orderItems));

        // --- Projeções por Recorrência Dinâmicas ---
        if (isRecurrenceActive) {
            const statusFilters = activeFilters.filter(f => ['buy', 'attention', 'suggest'].includes(f));
            
            let baseRecItems = [];
            let recMonths = 1;
            let recTitle = "📈 Projeções por Recorrência (1m)";
            let recColor = "#10b981";

            if (statusFilters.length === 1) {
                const currentFilter = statusFilters[0];
                if (currentFilter === 'buy') {
                    baseRecItems = buyItems;
                    recMonths = 1;
                    recTitle = "📈 Projeções por Recorrência (1m)";
                    recColor = "#fb7185";
                } else if (currentFilter === 'attention') {
                    baseRecItems = attentionItems;
                    recMonths = 2;
                    recTitle = "📈 Projeções por Recorrência (2m)";
                    recColor = "#f59e0b";
                } else if (currentFilter === 'suggest') {
                    baseRecItems = suggestItems;
                    recMonths = 3;
                    recTitle = "📈 Projeções por Recorrência (3m)";
                    recColor = "#eab308";
                }
            } else {
                baseRecItems = [...buyItems, ...attentionItems, ...suggestItems];
                recTitle = "📈 Projeções por Recorrência (Geral)";
                recColor = "#10b981";
            }

            const recTitleEl = document.getElementById('recurrence-title');
            if (recTitleEl) {
                recTitleEl.textContent = recTitle;
                recTitleEl.style.color = recColor;
            }

            const recSteps = [17, 33, 50, 67, 83, 100];
            const currentSliderIdx = document.getElementById('rec-slider')?.value || 0;
            const currentRecVal = recSteps[currentSliderIdx];
            
            const recDisplayEl = document.getElementById('current-rec-display');
            if (recDisplayEl) {
                recDisplayEl.textContent = `${currentRecVal}%`;
                recDisplayEl.style.color = recColor;
            }

            const updateRecSliderStats = (items) => {
                const countEl = document.getElementById('rec-stat-count');
                const valueEl = document.getElementById('rec-stat-value');
                const slider = document.getElementById('rec-slider');
                
                const filteredByRec = items.filter(i => parseFloat(i.recorrencia) >= currentRecVal);
                if (countEl) countEl.textContent = `${filteredByRec.length} itens`;
                
                let totalVal = 0;
                if (statusFilters.length === 1) {
                    totalVal = sumValAlerts(filteredByRec, recMonths);
                } else {
                    totalVal = filteredByRec.reduce((acc, i) => {
                        let m = 0;
                        if (i.situacao === 'ruptura') m = 1;
                        else if (i.situacao === 'atencao') m = 2;
                        else if (i.situacao === 'sugestao') m = 3;
                        return acc + ((parseFloat(i.medVenda) || 0) * 3 * m * (parseFloat(i.custoRaw) || 0));
                    }, 0);
                }
                
                if (valueEl) {
                    valueEl.textContent = formatValue(totalVal);
                    valueEl.style.color = recColor;
                }

                document.querySelectorAll('.timeline-legend strong').forEach(el => {
                    el.style.color = recColor;
                });

                if (slider) {
                    const style = document.createElement('style');
                    style.innerHTML = `
                        #rec-slider::-webkit-slider-thumb { background: ${recColor} !important; box-shadow: 0 0 15px ${recColor}66 !important; }
                        #rec-slider::-moz-range-thumb { background: ${recColor} !important; box-shadow: 0 0 15px ${recColor}66 !important; }
                    `;
                    document.head.appendChild(style);
                }
            };

            updateRecSliderStats(baseRecItems);
        }

        updateChart(data);
    }

    // --- File Handling ---

    /**
     * Renders a colored badge for Dias Úteis de Estoque
     * Green > 66 d.u. (>3 meses) | Amber 22-66 d.u. (1-3 meses) | Red < 22 d.u. (<1 mês)
     * Referência: ~22 dias úteis por mês no Brasil
     */
    function renderDiasEstoque(dias) {
        if (dias === null || dias === undefined) {
            return '<span style="font-size:0.72rem; color:var(--text-muted); background:rgba(255,255,255,0.06); padding:3px 8px; border-radius:6px;">S/Hist</span>';
        }
        let color, bg, tooltip;
        if (dias < 22) {
            color = '#fb7185'; bg = 'rgba(251,113,133,0.15)';
            tooltip = 'Cobertura crítica: menos de 1 mês de estoque (dias úteis)';
        } else if (dias <= 66) {
            color = '#f59e0b'; bg = 'rgba(245,158,11,0.15)';
            tooltip = 'Cobertura moderada: 1 a 3 meses de estoque (dias úteis)';
        } else {
            color = '#34d399'; bg = 'rgba(52,211,153,0.15)';
            tooltip = 'Cobertura saudável: mais de 3 meses de estoque (dias úteis)';
        }
        const label = dias > 999 ? '999+ d.u.' : `${dias} d.u.`;
        return `<span style="font-size:0.82rem; font-weight:700; color:${color}; background:${bg}; padding:3px 10px; border-radius:6px; border:1px solid ${color}33;" title="${tooltip}">${label}</span>`;
    }

    /**
     * Generates an enlarged SVG sparkline for history visualization
     */
    function generateSparkline(history, labels) {
        if (!history || history.length === 0) return '<span style="color:var(--text-muted); font-size:0.7rem;">Sem histórico</span>';
        
        const width = 750;
        const height = 180;
        const leftPadding = 40;
        const rightPadding = 55;  // extra espaço à direita para não cortar o último label
        const topPadding = 35;
        const bottomPadding = 32;
        
        const max = Math.max(...history, 5);
        const drawableWidth = width - leftPadding - rightPadding;
        const drawableHeight = height - topPadding - bottomPadding;
        
        const stepX = history.length > 1 ? drawableWidth / (history.length - 1) : 0;
        
        const points = history.map((val, i) => {
            const x = leftPadding + i * stepX;
            const y = height - bottomPadding - (val / max * drawableHeight);
            return { x, y, val, label: labels[i] };
        });

        const linePath = points.map((p, i) => (i === 0 ? 'M' : 'L') + `${p.x},${p.y}`).join(' ');
        const areaPath = linePath + ` L${points[points.length-1].x},${height - bottomPadding} L${points[0].x},${height - bottomPadding} Z`;

        return `
            <svg width="100%" height="${height}" viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMinYMin meet" style="overflow: visible; display: block; max-width: ${width}px;">
                <defs>
                    <linearGradient id="sparklineGradientDetail" x1="0%" y1="0%" x2="0%" y2="100%">
                        <stop offset="0%" style="stop-color:var(--info); stop-opacity:0.4" />
                        <stop offset="100%" style="stop-color:var(--info); stop-opacity:0.01" />
                    </linearGradient>
                </defs>
                <path d="${areaPath}" fill="url(#sparklineGradientDetail)" />
                <path d="${linePath}" fill="none" stroke="var(--info)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" />
                ${points.map(p => `
                    <circle cx="${p.x}" cy="${p.y}" r="5" fill="var(--info)" class="spark-dot">
                        <title>${p.label}: ${p.val}</title>
                    </circle>
                    <text x="${p.x}" y="${p.y - 14}" text-anchor="middle" font-size="12" font-weight="bold" fill="white" style="font-family: 'Outfit', sans-serif;">
                        ${p.val}
                    </text>
                    <text x="${p.x}" y="${height - 6}" text-anchor="middle" font-size="10.5" font-weight="600" fill="var(--text-muted)" style="font-family: 'Outfit', sans-serif;">
                        ${p.label}
                    </text>
                `).join('')}
            </svg>
        `;
    }

    function handleFile(file) {
        if (!file) return;

        showLoading();

        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheet];
                
                // Use defval: null to ensure all columns exist in all objects
                const json = XLSX.utils.sheet_to_json(worksheet, { 
                    defval: null,
                    blankrows: false 
                });

                console.log('JSON Bruto carregado:', json.length, 'linhas.');
                
                if (json.length === 0) {
                    alert('A planilha parece estar vazia ou mal formatada.');
                    return;
                }

                currentData = processData(json);
                initFixedCharts(currentData);
                
                showDashboard();
                applyAllFilters();
                
                // Pequeno delay para garantir que o DOM renderizou antes de tirar o loading
                setTimeout(hideLoading, 500);
            } catch (err) {
                console.error('Erro ao processar arquivo:', err);
                alert('Ocorreu um erro ao ler a planilha. Verifique se o formato está correto.');
                hideLoading();
            }

        };
        reader.readAsArrayBuffer(file);
    }

    /**
     * Handles Buyer Mapping File
     */
    function handleBuyerFile(file) {
        if (!file) return;
        showLoading();
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(worksheet);
                
                const newMap = {};
                json.forEach(row => {
                    const codeObj = getValue(row, ['Código', 'Cód.', 'Produto', 'Cod', 'ID', 'Item']);
                    const nameObj = getValue(row, ['Comprador', 'Nome', 'Responsável', 'Buyer', 'Nome do Comprador']);
                    
                    if (codeObj.value && nameObj.value) {
                        const code = codeObj.value.toString().trim();
                        const name = nameObj.value.toString().trim();
                        newMap[code] = name;
                        
                        // Also store localized version (without leading zeros or dots) for better matching
                        const cleanCode = code.replace(/^0+/, '').replace(/[.]/g, '');
                        if (cleanCode !== code) newMap[cleanCode] = name;
                    }
                });

                if (Object.keys(newMap).length === 0) {
                    alert('Nenhum mapeamento de "Código" e "Comprador" foi encontrado no arquivo.');
                    return;
                }

                buyerMap = newMap;
                localStorage.setItem('buyerMap', JSON.stringify(buyerMap));
                alert(`Sucesso! ${Object.keys(newMap).length} códigos vinculados. Agora você pode importar sua planilha de estoque.`);
                
                // If there's already data, re-process it to show buyers immediately
                // However, processData needs rawData which we don't store globally in its raw form.
                // For now, simple alert is enough; the user will re-upload or it's ready for next.
                hideLoading();
            } catch (err) {
                console.error('Erro ao processar arquivo de compradores:', err);
                alert('Erro ao ler arquivo de compradores.');
                hideLoading();
            }

        };
        reader.readAsArrayBuffer(file);
    }

    function showDashboard() {
        dropZone.style.display = 'none';
        chartSection.style.display = 'flex';
        statsSection.style.display = 'flex';
        document.getElementById('criteria-legend').style.display = 'flex';
        tableContainer.style.display = 'block';
    }

    function initFixedCharts(data) {
        const supplierStats = {};
        data.filter(i => i.situacao !== 'seguro').forEach(i => {
            supplierStats[i.fornecedor] = (supplierStats[i.fornecedor] || 0) + 1;
        });
        fixedSupplierLabels = Object.entries(supplierStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15)
            .map(s => s[0]);

        const groupStats = {};
        data.filter(i => i.situacao !== 'seguro').forEach(i => {
            groupStats[i.grupo] = (groupStats[i.grupo] || 0) + 1;
        });
        fixedGroupLabels = Object.entries(groupStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15)
            .map(g => g[0]);
    }

    function updateChart(data) {
        const buyCount = data.filter(i => i.situacao === 'ruptura').length;
        const attentionCount = data.filter(i => i.situacao === 'atencao').length;
        const suggestCount = data.filter(i => i.situacao === 'sugestao').length;
        const okCount = data.length - buyCount - attentionCount - suggestCount;

        const ctx = document.getElementById('purchase-pie-chart').getContext('2d');
        
        // Dynamic colors for doughnut slices based on active status filters
        const statusFilters = activeFilters.filter(f => ['buy', 'attention', 'suggest', 'safe'].includes(f));
        const baseColors = ['#fb7185', '#f59e0b', '#818cf8', '#34d399'];
        const filterKeys = ['buy', 'attention', 'suggest', 'safe'];
        const backgroundColors = baseColors.map((color, idx) => {
            const key = filterKeys[idx];
            if (statusFilters.length === 0) return color;
            return statusFilters.includes(key) ? color : color + '22'; // 22 is ~13% opacity in hex
        });

        const chartData = {
            labels: ['Ruptura', 'Atenção', 'Sugestão', 'Seguro'],
            datasets: [{
                data: [buyCount, attentionCount, suggestCount, okCount],
                backgroundColor: backgroundColors,
                hoverBackgroundColor: baseColors,
                borderColor: '#0f172a',
                borderWidth: 3,
                hoverOffset: 20,
                borderRadius: 4
            }]
        };

        if (myChart) {
            myChart.data = chartData;
            myChart.update();
        } else {
            myChart = new Chart(ctx, {
                type: 'doughnut',
                data: chartData,
                plugins: [
                    ChartDataLabels,
                    {
                        id: 'centerText',
                        beforeDraw: function(chart) {
                            const { width, height, ctx } = chart;
                            ctx.restore();
                            
                            // "TOTAL" label
                            ctx.font = `600 0.8rem Outfit, sans-serif`;
                            ctx.textBaseline = "middle";
                            ctx.fillStyle = "#9ca3af";
                            ctx.letterSpacing = "2px";
                            
                            const label = "TOTAL",
                                labelX = Math.round((width - ctx.measureText(label).width) / 2),
                                labelY = height / 2 - 22;
                            ctx.fillText(label, labelX, labelY);

                            // Value number
                            ctx.font = `700 2.2rem Outfit, sans-serif`;
                            ctx.fillStyle = "#ffffff";
                            ctx.letterSpacing = "0px";
                            const val = data.length.toString(),
                                valX = Math.round((width - ctx.measureText(val).width) / 2),
                                valY = height / 2 + 10;
                            ctx.fillText(val, valX, valY);
                            ctx.save();
                        }
                    }
                ],
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '78%',
                    spacing: 2,
                    layout: { padding: 10 },
                    onClick: (e) => {
                        const elements = myChart.getElementsAtEventForMode(e, 'nearest', { intersect: true }, false);
                        if (elements.length > 0) {
                            const index = elements[0].index;
                            const filters = ['buy', 'attention', 'suggest', 'safe'];
                            const filter = filters[index];
                            if (filter) {
                                isRecurrenceActive = (filter !== 'safe');
                                activeRecFilter = null;
                                document.querySelectorAll('.rec-card').forEach(c => c.style.boxShadow = 'none');
                                toggleFilter(filter, 'card');
                            }
                        }
                    },
                    onHover: (event, chartElement) => {
                        if (event && event.native && event.native.target) {
                            event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                        }
                    },
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: '#9ca3af',
                                usePointStyle: true,
                                pointStyle: 'circle',
                                padding: 15,
                                font: { family: 'Outfit', size: 11, weight: '500' }
                            }
                        },
                        title: {
                            display: true,
                            text: 'Status de Abastecimento',
                            color: '#ffffff',
                            padding: { bottom: 15 },
                            font: { family: 'Outfit', size: 16, weight: '700' }
                        },
                        datalabels: {
                            color: '#fff',
                            backgroundColor: 'rgba(15, 23, 42, 0.7)',
                            borderRadius: 4,
                            padding: { left: 6, right: 6, top: 4, bottom: 4 },
                            font: { weight: '700', size: 10 },
                            formatter: (value, ctx) => {
                                let sum = ctx.chart.data.datasets[0].data.reduce((acc, val) => acc + val, 0);
                                if (sum === 0) return '';
                                let percentage = (value * 100 / sum).toFixed(0) + "%";
                                return (value > 2) ? percentage : ''; // hide if too small
                            }
                        }
                    }
                }
            });
        }

        // --- Supplier Bar Chart ---
        const supplierCtx = document.getElementById('supplier-risk-chart').getContext('2d');
        
        const currentSuppStats = {};
        data.filter(i => {
            if (statusFilters.length > 0) {
                if (i.situacao === 'ruptura' && statusFilters.includes('buy')) return true;
                if (i.situacao === 'atencao' && statusFilters.includes('attention')) return true;
                if (i.situacao === 'sugestao' && statusFilters.includes('suggest')) return true;
                if (i.situacao === 'seguro' && statusFilters.includes('safe')) return true;
                return false;
            }
            return i.situacao !== 'seguro'; // Default: exclude seguro
        }).forEach(i => {
            currentSuppStats[i.fornecedor] = (currentSuppStats[i.fornecedor] || 0) + 1;
        });

        // Dynamic labels: top 15 suppliers for currently filtered view, sorted descending
        const dynamicSupplierLabels = Object.entries(currentSuppStats)
            .filter(entry => entry[1] > 0)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15)
            .map(s => s[0]);
        const supplierLabelsToShow = dynamicSupplierLabels.length > 0 ? dynamicSupplierLabels : fixedSupplierLabels;

        const supplierChartData = {
            labels: supplierLabelsToShow,
            datasets: [{
                label: 'Sugerido/Risco',
                data: supplierLabelsToShow.map(label => currentSuppStats[label] || 0),
                backgroundColor: supplierLabelsToShow.map(label => 
                    selectedSuppliers.includes(label) ? '#818cf8' : 'rgba(99, 102, 241, 0.4)'
                ),
                borderColor: '#6366f1',
                borderWidth: 2,
                borderRadius: 5,
            }]
        };

        if (supplierChart) {
            supplierChart.data = supplierChartData;
            supplierChart.update();
        } else {
            supplierChart = new Chart(supplierCtx, {
                type: 'bar',
                data: supplierChartData,
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (e) => {
                        const elements = supplierChart.getElementsAtEventForMode(e, 'nearest', { intersect: true }, false);
                        let label = null;

                        if (elements.length > 0) {
                            const index = elements[0].index;
                            label = supplierChart.data.labels[index];
                        } else {
                            // Se não clicou na barra, verifica se clicou no label (texto) à esquerda
                            const yAxis = supplierChart.scales.y;
                            if (e.x <= supplierChart.chartArea.left) {
                                const yValue = yAxis.getValueForPixel(e.y);
                                if (yValue >= 0 && yValue < supplierChart.data.labels.length) {
                                    const index = Math.round(yValue);
                                    label = supplierChart.data.labels[index];
                                }
                            }
                        }

                        if (label) {
                            if (selectedSuppliers.includes(label)) {
                                selectedSuppliers = selectedSuppliers.filter(s => s !== label);
                            } else {
                                selectedSuppliers.push(label);
                            }
                            applyAllFilters();
                        }
                    },
                    onHover: (event, chartElement) => {
                        if (event && event.native && event.native.target) {
                            event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        title: {
                            display: true,
                            text: 'Risco por Fornecedor (Filtrado)',
                            color: '#f3f4f6',
                            font: { family: 'Outfit', size: 18, weight: '600' }
                        },
                        datalabels: { display: false }
                    },
                    scales: {
                        x: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255, 255, 255, 0.05)' },
                            ticks: { color: '#9ca3af' }
                        },
                        y: {
                            grid: { display: false },
                            ticks: { 
                                color: '#9ca3af',
                                font: { size: 11 }
                            }
                        }
                    }
                }
            });
        }

        // --- Group Bar Chart (Horizontal) ---
        const groupCtx = document.getElementById('group-bar-chart').getContext('2d');
        
        const currentGroupStats = {};
        data.filter(i => {
            if (statusFilters.length > 0) {
                if (i.situacao === 'ruptura' && statusFilters.includes('buy')) return true;
                if (i.situacao === 'atencao' && statusFilters.includes('attention')) return true;
                if (i.situacao === 'sugestao' && statusFilters.includes('suggest')) return true;
                if (i.situacao === 'seguro' && statusFilters.includes('safe')) return true;
                return false;
            }
            return i.situacao !== 'seguro'; // Default: exclude seguro
        }).forEach(i => {
            currentGroupStats[i.grupo] = (currentGroupStats[i.grupo] || 0) + 1;
        });

        // Dynamic labels: top 15 groups for currently filtered view, sorted descending
        const dynamicGroupLabels = Object.entries(currentGroupStats)
            .filter(entry => entry[1] > 0)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15)
            .map(g => g[0]);
        const groupLabelsToShow = dynamicGroupLabels.length > 0 ? dynamicGroupLabels : fixedGroupLabels;

        const groupDataDetails = {
            labels: groupLabelsToShow,
            datasets: [{
                label: 'Sugerido/Risco',
                data: groupLabelsToShow.map(label => currentGroupStats[label] || 0),
                backgroundColor: groupLabelsToShow.map(label => 
                    selectedGroups.includes(label) ? '#f59e0b' : 'rgba(245, 158, 11, 0.4)'
                ),
                borderColor: '#f59e0b',
                borderWidth: 2,
                borderRadius: 5,
            }]
        };

        if (groupChart) {
            groupChart.data = groupDataDetails;
            groupChart.update();
        } else {
            groupChart = new Chart(groupCtx, {
                type: 'bar',
                data: groupDataDetails,
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (e) => {
                        const elements = groupChart.getElementsAtEventForMode(e, 'nearest', { intersect: true }, false);
                        let label = null;

                        if (elements.length > 0) {
                            const index = elements[0].index;
                            label = groupChart.data.labels[index];
                        } else {
                            const yAxis = groupChart.scales.y;
                            if (e.x <= groupChart.chartArea.left) {
                                const yValue = yAxis.getValueForPixel(e.y);
                                if (yValue >= 0 && yValue < groupChart.data.labels.length) {
                                    const index = Math.round(yValue);
                                    label = groupChart.data.labels[index];
                                }
                            }
                        }

                        if (label) {
                            if (selectedGroups.includes(label)) {
                                selectedGroups = selectedGroups.filter(g => g !== label);
                            } else {
                                selectedGroups.push(label);
                            }
                            applyAllFilters();
                        }
                    },
                    onHover: (event, chartElement) => {
                        if (event && event.native && event.native.target) {
                            event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        title: {
                            display: true,
                            text: 'Risco por Grupo (Filtrado)',
                            color: '#f3f4f6',
                            font: { family: 'Outfit', size: 18, weight: '600' }
                        },
                        datalabels: { display: false }
                    },
                    scales: {
                        x: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255, 255, 255, 0.05)' },
                            ticks: { color: '#9ca3af' }
                        },
                        y: {
                            grid: { display: false },
                            ticks: { color: '#9ca3af' }
                        }
                    }
                }
            });
        }
    }

    function updateRecurrenceVisibility() {
        const recurrenceGroup = document.getElementById('recurrence-group');
        if (!recurrenceGroup) return;

        // Só exibe se foi ativado via card E existe algum critério de projeção ativo
        const hasActiveCriterion = activeFilters.some(f => ['buy', 'attention', 'suggest'].includes(f));
        
        if (isRecurrenceActive && hasActiveCriterion) {
            recurrenceGroup.style.display = 'flex';
        } else {
            recurrenceGroup.style.display = 'none';
        }
    }

    function applyAllFilters() {
        updateRecurrenceVisibility();
        const prodTerm = tableSearch.value.toLowerCase();
        const smartTerm = smartSearch.value.toLowerCase().trim();
        const keywords = smartTerm.split(/\s+/).filter(k => k.length > 0);

        // 1. Context Filter (Buyer, Suppliers, Groups)
        const matchesContext = (i) => {
            const matchBuyer = activeBuyer === 'all' || i.comprador.toLowerCase() === activeBuyer.toLowerCase();
            const matchSelectedSupp = selectedSuppliers.length === 0 || selectedSuppliers.includes(i.fornecedor);
            const matchSelectedGroup = selectedGroups.length === 0 || selectedGroups.includes(i.grupo);
            return matchBuyer && matchSelectedSupp && matchSelectedGroup;
        };

        // 2. Search Filter (Table Search, Smart Search)
        const matchesSearch = (i) => {
            const cleanProd = i.produto.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            const cleanTerm = prodTerm.replace(/[^a-zA-Z0-9]/g, '');
            const matchProd = cleanProd.includes(cleanTerm) || 
                               i.grupo.toString().toLowerCase().includes(prodTerm) ||
                               i.descricao.toString().toLowerCase().includes(prodTerm);
            
            const matchSmart = keywords.every(kw => 
                i.descricao.toLowerCase().includes(kw) || 
                i.produto.toString().toLowerCase().includes(kw) ||
                i.fornecedor.toLowerCase().includes(kw)
            );
            return matchProd && matchSmart;
        };

        // 3. Status/Recurrence Filter
        const matchesStatus = (i) => {
            let matchRec = activeRecFilter === null || parseFloat(i.recorrencia) >= activeRecFilter;
            
            let matchStatus = true;
            if (activeFilters.includes('all')) {
                // PERFORMANCE OPTIMIZATION: By default, exclude "Safe" items from table rendering
                // They will only show if the 'safe' card is explicitly clicked.
                matchStatus = (i.situacao !== 'seguro');
            } else {
                const statusFilters = activeFilters.filter(f => ['buy', 'attention', 'suggest', 'safe'].includes(f));
                if (statusFilters.length > 0) {
                    matchStatus = false;
                    if (statusFilters.includes('buy') && i.situacao === 'ruptura') matchStatus = true;
                    if (statusFilters.includes('attention') && i.situacao === 'atencao') matchStatus = true;
                    if (statusFilters.includes('suggest') && i.situacao === 'sugestao') matchStatus = true;
                    if (statusFilters.includes('safe') && i.situacao === 'seguro') matchStatus = true;
                }
            }

            if (activeFilters.includes('has-order') && !i.temEncomenda) return false;
            return matchStatus && matchRec;
        };

        // Data for Stats and Charts: Context Only
        const chartData = currentData.filter(i => matchesContext(i));

        // Data for Summary Cards: Context + Search + Recurrence
        const summaryData = chartData.filter(i => {
            const matchSearch = matchesSearch(i);
            const matchRec = activeRecFilter === null || parseFloat(i.recorrencia) >= activeRecFilter;
            return matchSearch && matchRec;
        });

        // Data for Table: Full Filtering (Includes Status Toggle)
        filteredData = chartData.filter(i => matchesSearch(i) && matchesStatus(i));

        renderActiveFilters();

        // 4. Sorting
        if (sortVendasDir === 'desc') {
            filteredData.sort((a, b) => parseFloat(b.vendas) - parseFloat(a.vendas));
        } else if (sortVendasDir === 'asc') {
            filteredData.sort((a, b) => parseFloat(a.vendas) - parseFloat(b.vendas));
        } else if (sortRecorrenciaDir === 'desc') {
            filteredData.sort((a, b) => parseFloat(b.recorrencia) - parseFloat(a.recorrencia));
        } else if (sortRecorrenciaDir === 'asc') {
            filteredData.sort((a, b) => parseFloat(a.recorrencia) - parseFloat(b.recorrencia));
        } else if (sortDiasEstoqueDir === 'asc') {
            filteredData.sort((a, b) => (a.diasEstoque === null ? 1 : b.diasEstoque === null ? -1 : a.diasEstoque - b.diasEstoque));
        } else if (sortDiasEstoqueDir === 'desc') {
            filteredData.sort((a, b) => (a.diasEstoque === null ? 1 : b.diasEstoque === null ? -1 : b.diasEstoque - a.diasEstoque));
        }

        updateDashboardUI(summaryData);
        renderTable(filteredData);
    }

    function renderActiveFilters() {
        const container = document.getElementById('active-filters-container');
        if (!container) return;
        container.innerHTML = '';

        if (selectedSuppliers.length === 0 && selectedGroups.length === 0 && activeFilters.includes('all') && activeBuyer === 'all' && activeRecFilter === null) {
            container.innerHTML = '<span class="no-filters">Nenhum filtro ativo</span>';
            return;
        }
        
        container.style.display = 'flex';
        
        selectedSuppliers.forEach(supp => {
            const tag = document.createElement('div');
            tag.className = 'filter-tag';
            tag.innerHTML = `
                <span>🏢 ${supp}</span>
                <span class="filter-tag-remove" data-type="supplier" data-value="${supp}">&times;</span>
            `;
            container.appendChild(tag);
        });
        
        selectedGroups.forEach(group => {
            const tag = document.createElement('div');
            tag.className = 'filter-tag group-tag';
            tag.innerHTML = `
                <span>📦 ${group}</span>
                <span class="filter-tag-remove" data-type="group" data-value="${group}">&times;</span>
            `;
            container.appendChild(tag);
        });
        
        // Add event listeners for removal
        container.querySelectorAll('.filter-tag-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const type = btn.getAttribute('data-type');
                const value = btn.getAttribute('data-value');
                
                if (type === 'supplier') {
                    selectedSuppliers = selectedSuppliers.filter(s => s !== value);
                } else if (type === 'group') {
                    selectedGroups = selectedGroups.filter(g => g !== value);
                }
                
                applyAllFilters();
            });
        });
    }

    function toggleFilter(filter, source = 'table') {
        if (filter === 'all') {
            activeFilters = ['all'];
            isRecurrenceActive = false;
        } else {
            // Se o clique veio dos botões da tabela ou outro lugar que não seja o card resumo
            if (source === 'table') {
                isRecurrenceActive = false;
            }

            // Remove 'all' if it exists
            activeFilters = activeFilters.filter(f => f !== 'all');
            
            if (activeFilters.includes(filter)) {
                // Toggle off
                activeFilters = activeFilters.filter(f => f !== filter);
                if (activeFilters.length === 0) isRecurrenceActive = false;
            } else {
                // Toggle on
                activeFilters.push(filter);
            }

            // If nothing is selected, default to 'all'
            if (activeFilters.length === 0) {
                activeFilters = ['all'];
                isRecurrenceActive = false;
            }
        }

        // Update UI - Table Buttons
        filterBtns.forEach(btn => {
            const btnFilter = btn.getAttribute('data-filter');
            if (activeFilters.includes(btnFilter)) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        // Update UI - Summary Cards
        document.querySelectorAll('.stat-card').forEach(card => {
            const cardFilter = card.getAttribute('data-filter');
            if (cardFilter && activeFilters.includes(cardFilter)) {
                card.classList.add('selected');
            } else {
                card.classList.remove('selected');
            }
        });

        applyAllFilters();
    }

    // --- Events ---

    fileUpload.addEventListener('change', (e) => {
        handleFile(e.target.files[0]);
    });

    // Combined search logic
    tableSearch.addEventListener('input', applyAllFilters);
    smartSearch.addEventListener('input', applyAllFilters);
    
    buyerBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            activeBuyer = btn.getAttribute('data-buyer');
            buyerBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            applyAllFilters();
        });
    });

    buyerUpload.addEventListener('change', (e) => {
        handleBuyerFile(e.target.files[0]);
    });

    // Sorting functionality
    const sortRecHeader = document.getElementById('sort-recorrencia');
    const sortVendasHeader = document.getElementById('sort-vendas');
    const sortDiasHeader = document.getElementById('sort-dias-estoque');

    function resetSortIcons(...excludes) {
        const allHeaders = [sortRecHeader, sortVendasHeader, sortDiasHeader];
        allHeaders.forEach(h => {
            if (h && !excludes.includes(h)) h.querySelector('.sort-icon').textContent = '↕️';
        });
    }

    if (sortRecHeader) {
        sortRecHeader.addEventListener('click', () => {
            sortVendasDir = 'none';
            sortDiasEstoqueDir = 'none';
            resetSortIcons(sortRecHeader);
            if (sortRecorrenciaDir === 'none' || sortRecorrenciaDir === 'asc') {
                sortRecorrenciaDir = 'desc';
                sortRecHeader.querySelector('.sort-icon').textContent = '🔽';
            } else {
                sortRecorrenciaDir = 'asc';
                sortRecHeader.querySelector('.sort-icon').textContent = '🔼';
            }
            applyAllFilters();
        });
    }

    if (sortVendasHeader) {
        sortVendasHeader.addEventListener('click', () => {
            sortRecorrenciaDir = 'none';
            sortDiasEstoqueDir = 'none';
            resetSortIcons(sortVendasHeader);
            if (sortVendasDir === 'none' || sortVendasDir === 'asc') {
                sortVendasDir = 'desc';
                sortVendasHeader.querySelector('.sort-icon').textContent = '🔽';
            } else {
                sortVendasDir = 'asc';
                sortVendasHeader.querySelector('.sort-icon').textContent = '🔼';
            }
            applyAllFilters();
        });
    }

    if (sortDiasHeader) {
        sortDiasHeader.addEventListener('click', () => {
            sortRecorrenciaDir = 'none';
            sortVendasDir = 'none';
            resetSortIcons(sortDiasHeader);
            if (sortDiasEstoqueDir === 'none' || sortDiasEstoqueDir === 'desc') {
                sortDiasEstoqueDir = 'asc'; // menor dias primeiro = mais crítico
                sortDiasHeader.querySelector('.sort-icon').textContent = '🔼';
            } else {
                sortDiasEstoqueDir = 'desc';
                sortDiasHeader.querySelector('.sort-icon').textContent = '🔽';
            }
            applyAllFilters();
        });
    }

    // Multi-select status filters
    filterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const filter = btn.getAttribute('data-filter');
            // Clear recurrence filter when switching status if needed, 
            // but let's keep them independent for power users.
            toggleFilter(filter);
        });
    });

    // Clickable Projection Cards (Ruptura, Atenção, Sugestão)
    document.querySelectorAll('.clickable-card').forEach(card => {
        card.addEventListener('click', () => {
            const filter = card.getAttribute('data-filter');
            isRecurrenceActive = true; // Ativa a exibição da seção de recorrência
            activeRecFilter = null; // Reset rec filter when clicking main status
            document.querySelectorAll('.rec-card').forEach(c => c.style.boxShadow = 'none');
            toggleFilter(filter, 'card');
        });
    });

    // Recurrence Slider (Timeline) Event
    const recSlider = document.getElementById('rec-slider');
    if (recSlider) {
        recSlider.addEventListener('input', (e) => {
            const recSteps = [17, 33, 50, 67, 83, 100];
            const val = recSteps[e.target.value];
            activeRecFilter = val;
            
            // Re-render table and stats (this will call updateRecSliderStats via renderTable)
            applyAllFilters();
        });
    }

    // Clear filters
    clearFiltersBtn.addEventListener('click', () => {
        tableSearch.value = '';
        smartSearch.value = '';
        selectedSuppliers = [];
        selectedGroups = [];
        activeBuyer = 'all';
        activeRecFilter = null;
        isRecurrenceActive = false;
        if (recSlider) recSlider.value = 0; // Reset to 17%
        
        buyerBtns.forEach(btn => {
            if (btn.getAttribute('data-buyer') === 'all') btn.classList.add('active');
            else btn.classList.remove('active');
        });
        activeFilters = ['all'];
        filterBtns.forEach(btn => {
            const btnFilter = btn.getAttribute('data-filter');
            if (btnFilter === 'all') btn.classList.add('active');
            else btn.classList.remove('active');
        });
        
        document.querySelectorAll('.stat-card').forEach(card => {
            card.classList.remove('selected');
        });

        applyAllFilters();
    });

    // Unified Export Button
    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', () => exportToExcel());
    }

    // --- Export Logic ---
    window.exportToExcel = () => {
        if (filteredData.length === 0) {
            alert("Não há dados para exportar. Importe uma planilha primeiro.");
            return;
        }
        
        const todayLocale = new Date().toLocaleDateString('pt-BR');
        const sheetName = 'Lista Filtrada Exportada';

        // Prepare data with the specific columns requested by the user
        const exportData = filteredData.map(i => {
            // Determine multiplier based on item status
            let multiplier = 1;
            if (i.situacao === 'ruptura') multiplier = 1;
            else if (i.situacao === 'atencao') multiplier = 2;
            else if (i.situacao === 'sugestao') multiplier = 3;
            else multiplier = 0; // If it's safe, maybe suggest 0? 

            const qtdSugerida = Math.ceil(parseFloat(i.medVenda) * multiplier);
            
            return {
                "Produto": i.produto,
                "Quantidade": qtdSugerida,
                "Centro de Custo": 1,
                "Finalidade": 1,
                "Data da Necessidade": todayLocale
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

        // Format column widths for better readability (matching the 5 requested columns)
        const wscols = [
            {wch: 15}, // Produto
            {wch: 12}, // Quantidade
            {wch: 15}, // Centro de Custo
            {wch: 12}, // Finalidade
            {wch: 18}  // Data Necessidade
        ];
        worksheet['!cols'] = wscols;

        XLSX.writeFile(workbook, `exportacao_estoque_${todayLocale.replace(/\//g, '-')}.xlsx`);
    };

    window.exportAttention = () => exportToExcel(); // Kept for legacy if called elsewhere

    // --- Multi-Supplier Modal Logic ---
    if (openSupplierModalBtn && supplierModal) {
        openSupplierModalBtn.addEventListener('click', () => {
            renderSupplierChecklist();
            supplierModal.classList.add('active');
        });

        closeSupplierModalBtn.onclick = () => supplierModal.classList.remove('active');
        
        window.addEventListener('click', (e) => {
            if (e.target === supplierModal) supplierModal.classList.remove('active');
        });

        modalSupplierSearch.oninput = (e) => {
            const val = e.target.value.toLowerCase();
            const items = supplierChecklist.querySelectorAll('.check-item');
            items.forEach(item => {
                const text = item.innerText.toLowerCase();
                item.style.display = text.includes(val) ? 'flex' : 'none';
            });
        };

        btnSelectAllSuppliers.onclick = () => {
            const checkboxes = supplierChecklist.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => {
                if (cb.parentElement.style.display !== 'none') cb.checked = true;
            });
        };

        btnClearAllSuppliers.onclick = () => {
            const checkboxes = supplierChecklist.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => {
                if (cb.parentElement.style.display !== 'none') cb.checked = false;
            });
        };

        applySupplierFilterBtn.onclick = () => {
            const checkboxes = supplierChecklist.querySelectorAll('input[type="checkbox"]');
            selectedSuppliers = [];
            checkboxes.forEach(cb => {
                if (cb.checked) selectedSuppliers.push(cb.value);
            });
            supplierModal.classList.remove('active');
            applyAllFilters();
        };
    }

    function renderSupplierChecklist() {
        if (!supplierChecklist) return;
        const uniqueSuppliers = [...new Set(currentData.map(i => i.fornecedor))].sort();
        
        supplierChecklist.innerHTML = uniqueSuppliers.map(supp => `
            <label class="check-item" style="display: flex; align-items: center; gap: 10px; padding: 8px; border-radius: 6px; cursor: pointer; transition: background 0.2s;">
                <input type="checkbox" value="${supp}" ${selectedSuppliers.includes(supp) ? 'checked' : ''} style="width: 16px; height: 16px; accent-color: var(--primary);">
                <span style="font-size: 13px; color: #fff;">${supp}</span>
            </label>
        `).join('');
    }

});
