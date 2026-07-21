/**
 * gerar_dados.js
 * Script standalone para gerar data-entradas.js a partir dos XLSXs do OneDrive.
 * Uso: node gerar_dados.js
 */
const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const SOURCE_DIR  = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Compras\\Comprasbrasil - Compras\\Planilhas Gestor\\Análise de Entradas LFI 2.1.K";
const OUTPUT_FILE = path.join(__dirname, "data-entradas.js");

async function main() {
    console.log('\n GERANDO data-entradas.js');
    console.log('----------------------------------');
    console.log('Pasta: ' + SOURCE_DIR);
    console.log('Destino: ' + OUTPUT_FILE);
    console.log('----------------------------------\n');

    if (!fs.existsSync(SOURCE_DIR)) {
        console.error('ERRO: Diretorio nao encontrado: ' + SOURCE_DIR);
        process.exit(1);
    }

    const files = fs.readdirSync(SOURCE_DIR).filter(function(f) {
        return f.match(/\.(xlsx|xls|XLSX)$/);
    });
    console.log('Arquivos encontrados: ' + files.length);

    let masterData = [];

    for (const file of files) {
        const fullPath = path.join(SOURCE_DIR, file);
        console.log('  Processando: ' + file);
        try {
            const workbook = XLSX.readFile(fullPath, { cellDates: true });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(firstSheet);
            const enriched = rows.map(function(r) { 
                r._source_file = file; 
                return r; 
            });
            masterData = masterData.concat(enriched);
        } catch (e) {
            console.error('  AVISO - Erro ao processar ' + file + ': ' + e.message);
        }
    }

    console.log('\nTotal de registros: ' + masterData.length);

    const jsContent = '// Gerado automaticamente em ' + new Date().toISOString() + '\n// Fonte: Analise de Entradas LFI 2.1.K\nconst PRE_LOADED_ENTRADAS = ' + JSON.stringify(masterData, null, 2) + ';\n';

    fs.writeFileSync(OUTPUT_FILE, jsContent, 'utf8');
    
    const sizeKB = Math.round(fs.statSync(OUTPUT_FILE).size / 1024);
    console.log('data-entradas.js gerado com sucesso! (' + sizeKB + ' KB)');
    console.log('\nProximo passo:');
    console.log('   git add data-entradas.js .gitignore');
    console.log('   git commit -m "sync: atualiza dados de entradas"');
    console.log('   git push\n');
}

main().catch(function(e) {
    console.error('Erro fatal: ' + e.message);
    process.exit(1);
});