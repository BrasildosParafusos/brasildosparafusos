const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const PORT = 3000;
const SOURCE_DIR = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Compras\\Comprasbrasil - Compras\\Planilhas Gestor\\Análise de Entradas LFI 2.1.K";
const OUTPUT_FILE = "data-entradas.js";
const LEADTIME_FILE = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Compras\\Relatório Lead time\\Base Leadtime.xlsx";
const RUPTURE_DIR = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Compras\\Comprasbrasil - Compras\\Planilhas Gestor\\Análise Evolutiva EST 3.2.T";

/**
 * Normaliza nomes de colunas para busca flexível
 */
function normalizeKey(key) {
    if (!key) return "";
    return key.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function findColumn(row, possibilities) {
    const keys = Object.keys(row);
    for (const p of possibilities) {
        const normP = normalizeKey(p);
        const match = keys.find(k => normalizeKey(k) === normP || normalizeKey(k).includes(normP));
        if (match) return match;
    }
    return null;
}

function getValueInRow(row, possibilities) {
    const col = findColumn(row, possibilities);
    return col ? row[col] : null;
}

const server = http.createServer((req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }
    
    if (req.url === '/sync') {
        console.log("Recebido pedido de sincronização...");
        try {
            if (!fs.existsSync(SOURCE_DIR)) {
                throw new Error(`Diretório não encontrado: ${SOURCE_DIR}`);
            }

            const files = fs.readdirSync(SOURCE_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
            console.log(`Arquivos encontrados: ${files.length}`);
            
            let masterData = [];

            for (const file of files) {
                const fullPath = path.join(SOURCE_DIR, file);
                console.log(`Processando: ${file}`);
                
                const workbook = XLSX.readFile(fullPath, { cellDates: true });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                if (rows.length < 2) continue;

                // Encontrar a linha de cabeçalho (primeira que tenha 'data' ou 'movto')
                let headerIndex = -1;
                for (let i = 0; i < Math.min(rows.length, 15); i++) {
                    const rowStr = JSON.stringify(rows[i]).toLowerCase();
                    if (rowStr.includes('data') || rowStr.includes('movto') || rowStr.includes('cod') || rowStr.includes('nf')) {
                        headerIndex = i;
                        break;
                    }
                }

                if (headerIndex === -1) headerIndex = 0;

                const normalize = (s) => String(s || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                const headers = rows[headerIndex].map(h => normalize(h));
                const findIdx = (names) => headers.findIndex(h => names.some(n => h.includes(normalize(n))));

                const dIdx = findIdx(['dt.movto', 'data', 'movimento']);
                const cIdx = findIdx(['codigo', 'cod']);
                const desIdx = findIdx(['descricao', 'item', 'produto']);
                const qIdx = findIdx(['quantidade', 'qtd', 'unidades']);
                const nfIdx = findIdx(['nota fiscal', 'nf', 'doc']);
                const fIdx = findIdx(['finalid.item ordem com', 'finalidade']);
                const vIdx = findIdx(['vlr.cont.p/sped', 'vlr.cont', 'valor', 'total']);
                const sIdx = findIdx(['razao social', 'fornecedor', 'nome']);
                const gIdx = 11; // Coluna L (12ª coluna)

                const processed = rows.slice(headerIndex + 1).map(row => {
                    if (!row || row.length === 0) return null;
                    if (!row[vIdx] && !row[qIdx] && !row[cIdx]) return null;

                    let dateVal = row[dIdx];
                    let dateStr = "";

                    if (dateVal instanceof Date) {
                        dateStr = dateVal.toISOString().split('T')[0];
                    } else if (typeof dateVal === 'number') {
                        const d = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
                        dateStr = d.toISOString().split('T')[0];
                    } else if (typeof dateVal === 'string') {
                        const parts = dateVal.split(/[\/\-]/);
                        if (parts.length === 3) {
                            if (parts[0].length === 4) dateStr = dateVal;
                            else dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
                        }
                    }

                    return {
                        date: dateStr,
                        code: String(row[cIdx] || '').trim(),
                        description: String(row[desIdx] || '').trim(),
                        quantity: parseFloat(row[qIdx]) || 0,
                        group: String(row[gIdx] || 'DIVERSOS').trim(),
                        invoice: String(row[nfIdx] || '').trim(),
                        purpose: String(row[fIdx] || '').trim(),
                        supplier: String(row[sIdx] || 'NÃO IDENTIFICADO').trim(),
                        value: parseFloat(row[vIdx]) || 0
                    };
                }).filter(row => row && (row.quantity !== 0 || row.value !== 0));

                masterData = masterData.concat(processed);
            }

            const jsContent = `const PRE_LOADED_ENTRADAS = ${JSON.stringify(masterData, null, 4)};`;
            fs.writeFileSync(OUTPUT_FILE, jsContent);
            
            console.log(`Sincronização concluída! ${masterData.length} registros totais.`);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ 
                success: true, 
                count: masterData.length,
                data: masterData 
            }));
        } catch (e) {
            console.error("ERRO NA SINCRONIZAÇÃO:", e.message);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: false, error: e.message }));
        }
    } else if (req.url === '/sales-history') {
        console.log("Recebido pedido de histórico de vendas (Leadtime)...");
        try {
            if (!fs.existsSync(LEADTIME_FILE)) {
                throw new Error(`Arquivo Leadtime não encontrado: ${LEADTIME_FILE}`);
            }

            const workbook = XLSX.readFile(LEADTIME_FILE);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Usar header: 1 para ler como array de arrays e identificar o cabeçalho corretamente
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (rows.length === 0) {
                throw new Error("Arquivo Leadtime vazio.");
            }

            // Identificar a linha de cabeçalho e as colunas de meses
            let headerRow = rows[0];
            // Se a primeira linha não parecer ter meses, procurar nas próximas 5
            for (let i = 0; i < Math.min(rows.length, 5); i++) {
                if (rows[i].some(cell => typeof cell === 'string' && /^\d{2}\/\d{4}$/.test(cell))) {
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
                    // EXCLUSÃO TOTAL: Somente até Maio/2026
                    return (yr < 2026) || (yr === 2026 && mon <= 5);
                })
                .sort((a, b) => {
                    const [mA, yA] = a.split('/').map(Number);
                    const [mB, yB] = b.split('/').map(Number);
                    return (yA * 12 + mA) - (yB * 12 + mB);
                });

            // Remover duplicatas caso o map tenha gerado labels iguais
            const uniqueMonthCols = [...new Set(monthCols)];

            console.log("Meses detectados no Leadtime (Final):", uniqueMonthCols);

            // Mapear índices das colunas de meses originais para os labels normalizados
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

            // Ordenar monthMap para seguir uniqueMonthCols
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

            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ 
                success: true, 
                count: Object.keys(salesMap).length,
                data: salesMap 
            }));
            console.log(`Histórico de vendas enviado: ${Object.keys(salesMap).length} itens.`);

        } catch (e) {
            console.error("ERRO AO LER HISTÓRICO DE VENDAS:", e.message);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: false, error: e.message }));
        }
    } else if (req.url === '/rupture-analysis') {
        console.log("Recebido pedido de análise profunda de rupturas...");
        try {
            if (!fs.existsSync(RUPTURE_DIR)) {
                fs.mkdirSync(RUPTURE_DIR, { recursive: true });
            }

            const files = fs.readdirSync(RUPTURE_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
            console.log(`Arquivos de ruptura encontrados: ${files.length}`);
            
            const historyData = [];

            for (const file of files) {
                const fullPath = path.join(RUPTURE_DIR, file);
                const stats = fs.statSync(fullPath);
                
                // Tenta extrair data do nome do arquivo (ex: 2024-05-07.xlsx) ou usa data de criação
                let fileDate = stats.mtime.toISOString().split('T')[0];
                const dateMatch = file.match(/(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{4})/);
                if (dateMatch) {
                    fileDate = dateMatch[0];
                }

                const workbook = XLSX.readFile(fullPath, { cellDates: true });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet);

                let ruptureCount = 0;
                let ruptureValue = 0;
                let ruptureInvest = 0;
                let attentionCount = 0;
                let attentionValue = 0;
                let attentionInvest = 0;
                let suggestCount = 0;
                let suggestValue = 0;
                let suggestInvest = 0;
                let totalItems = rows.length;

                // Identificar colunas de meses para cálculo de recorrência
                // Lê a primeira linha bruta para pegar os nomes das colunas (mesmo que sejam datas)
                const rawRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                const headerRow = rawRows[0] || [];
                
                const monthInfo = headerRow.map((val, idx) => {
                    let label = "";
                    if (val instanceof Date) {
                        const m = (val.getMonth() + 1).toString().padStart(2, '0');
                        const y = val.getFullYear();
                        label = `${m}/${y}`;
                    } else {
                        label = String(val || '');
                    }
                    return { label, index: idx, original: val };
                }).filter(h => /^\d{2}\/\d{4}$/.test(h.label));

                // Filtrar apenas meses até 05/2026 (mesmo do dashboard)
                const validMonthInfo = monthInfo.filter(h => {
                    const [m, y] = h.label.split('/').map(Number);
                    return (y < 2026) || (y === 2026 && m <= 5);
                });

                rows.forEach(row => {
                    const estoque = parseFloat(getValueInRow(row, ['estoque', 'saldo', 'atual'])) || 0;
                    const encomendas = parseFloat(getValueInRow(row, ['encomendas', 'pedido', 'transito', 'receber'])) || 0;
                    const custo = parseFloat(getValueInRow(row, ['preco reposicao', 'custo', 'unitario'])) || 0;
                    
                    // Cálculo de Vendas, Recorrência e Média idêntico ao app.js
                    const vendasRaw = parseFloat(getValueInRow(row, ['vendas', 'qtd. vendida', 'venda total', 'total vendas'])) || 0;
                    let totalVendas = vendasRaw;
                    let activeMonths = 0;
                    let recorrencia = 0;
                    let medVenda = 0;

                    // Identifica meses ativos para cálculo de recorrência
                    if (validMonthInfo.length > 0) {
                        let sumMonths = 0;
                        validMonthInfo.forEach(m => {
                            const val = parseFloat(row[m.label] || row[m.original] || 0);
                            if (val > 0) {
                                sumMonths += val;
                                activeMonths++;
                            }
                        });
                        
                        // Fallback: Se a coluna 'Vendas' for zero, usa a soma dos meses (mesmo do app.js)
                        if (totalVendas === 0) totalVendas = sumMonths;
                        
                        recorrencia = activeMonths / validMonthInfo.length;
                        // Média baseada apenas nos meses com venda (mesma lógica do app.js)
                        medVenda = activeMonths > 0 ? (totalVendas / activeMonths) : 0;
                    } else {
                        // ... fallback anterior
                        // Fallback caso não haja colunas de meses (usa as colunas agregadas)
                        medVenda = parseFloat(getValueInRow(row, ['med.venda', 'media', 'giro', 'venda mensal'])) || 0;
                        const recCol = findColumn(row, ['recorrencia', 'giro freq', 'frequencia']);
                        if (recCol) {
                            recorrencia = parseFloat(row[recCol]) || 0;
                            if (recorrencia > 1) recorrencia = recorrencia / 100;
                        } else {
                            recorrencia = 1;
                        }
                    }
                    
                    const totalDisponivel = estoque + encomendas;
                    const passesRecurrence = (recorrencia > 0.33);

                    if (passesRecurrence) {
                        if (medVenda > totalDisponivel) {
                            ruptureCount++;
                            ruptureValue += (medVenda * custo);
                            ruptureInvest += Math.max(0, (medVenda - totalDisponivel) * custo);
                        } else if ((medVenda * 2) > totalDisponivel) {
                            attentionCount++;
                            attentionValue += (medVenda * 2 * custo);
                            attentionInvest += Math.max(0, (medVenda * 2 - totalDisponivel) * custo);
                        } else if ((medVenda * 3) > totalDisponivel) {
                            suggestCount++;
                            suggestValue += (medVenda * 3 * custo);
                            suggestInvest += Math.max(0, (medVenda * 3 - totalDisponivel) * custo);
                        }
                    }
                });

                historyData.push({
                    file: file,
                    date: fileDate,
                    totalItems,
                    rupture: { count: ruptureCount, value: ruptureValue, invest: ruptureInvest },
                    attention: { count: attentionCount, value: attentionValue, invest: attentionInvest },
                    suggest: { count: suggestCount, value: suggestValue, invest: suggestInvest }
                });
            }

            // Ordenar por data
            historyData.sort((a, b) => new Date(a.date) - new Date(b.date));

            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ 
                success: true, 
                count: historyData.length,
                data: historyData 
            }));
            
        } catch (e) {
            console.error("ERRO NA ANÁLISE DE RUPTURAS:", e.message);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: false, error: e.message }));
        }
    } else {
        res.writeHead(404);
        res.end();
    }
});

server.listen(PORT, () => {
    console.log(`\n🚀 PONTE DE DADOS ATIVA`);
    console.log(`----------------------------------`);
    console.log(`URL: http://localhost:${PORT}/sync`);
    console.log(`Pasta: ${SOURCE_DIR}`);
    console.log(`Destino: ${path.resolve(OUTPUT_FILE)}`);
    console.log(`----------------------------------\n`);
});
