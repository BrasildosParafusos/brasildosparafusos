@echo off
echo ==========================================
echo  ATUALIZANDO DADOS DO GESTOR DE ESTOQUE
echo ==========================================
echo.

:: Passo 1: Gerar data-entradas.js a partir dos XLSXs do OneDrive
echo [1/3] Gerando dados de entradas...
node gerar_dados.js

if %ERRORLEVEL% NEQ 0 (
    echo ERRO ao gerar dados. Verifique se o Node.js esta instalado.
    pause
    exit /b 1
)

:: Passo 2: Commit e Push no GitHub
echo.
echo [2/3] Enviando para o GitHub...
git add data-entradas.js .gitignore gerar_dados.js bridge.js atualizar_dados.bat
git commit -m "sync: atualiza dados de entradas %date%"
git push origin main

:: Passo 3: Fazer upload para o servidor de revisao (Hostinger)
echo.
echo [3/3] Publicando no servidor de revisao...
curl -s -X POST "https://revisao.brasildosparafusos.com.br/compras/analise/deploy.php" ^
     -F "token=BrasildosParafusos2026!deploy" ^
     -F "data_file=@data-entradas.js"

echo.
echo ==========================================
echo  CONCLUIDO!
echo  O servidor de revisao foi atualizado.
echo  Acesse: https://revisao.brasildosparafusos.com.br/compras/analise/
echo ==========================================
echo.
pause