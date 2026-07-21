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

:: Passo 2: Commit no Git
echo.
echo [2/3] Commitando no Git...
git add data-entradas.js .gitignore gerar_dados.js bridge.js
git commit -m "sync: atualiza dados de entradas %date% %time:~0,5%"

:: Passo 3: Push para o GitHub (servidor de revisao vai precisar de deploy manual)
echo.
echo [3/3] Enviando para o GitHub...
git push origin main

echo.
echo ==========================================
echo  CONCLUIDO! 
echo  Agora acesse o servidor Hostinger e
echo  atualize o arquivo data-entradas.js
echo  via FTP ou File Manager.
echo ==========================================
echo.
pause