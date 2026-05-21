@echo off
setlocal
echo ==========================================
echo SINCRO-DASHBOARD: OneDrive -> XAMPP
echo ==========================================
echo.

set SOURCE=%~dp0
set DEST=c:\xampp\htdocs\análiseclasseabc

echo Origem:  %SOURCE%
echo Destino: %DEST%
echo.

if not exist "%DEST%" (
    echo [ERRO] A pasta de destino nao existe. Criando...
    mkdir "%DEST%"
)

echo Copiando arquivos...
xcopy "%SOURCE%index.html" "%DEST%\" /Y /Q
xcopy "%SOURCE%styles.css" "%DEST%\" /Y /Q
xcopy "%SOURCE%app.js" "%DEST%\" /Y /Q
xcopy "%SOURCE%logo-parafuso.png" "%DEST%\" /Y /Q
xcopy "%SOURCE%bridge.js" "%DEST%\" /Y /Q
xcopy "%SOURCE%rupturas.*" "%DEST%\" /Y /Q
xcopy "%SOURCE%entradas.*" "%DEST%\" /Y /Q
xcopy "%SOURCE%giro.*" "%DEST%\" /Y /Q

echo.
echo [SUCESSO] Arquivos atualizados no XAMPP!
echo Agora voce pode dar F5 no navegador.
echo.
pause
