@echo off
title Atualizar historico SISCON
color 0B
echo.
echo ============================================
echo   Registrando alteracoes e enviando ao GitHub
echo ============================================
echo.
echo Antes de rodar: coloque a planilha nova nesta pasta com o nome Solicitacoes.xlsx
echo.
if not exist "Solicitacoes.xlsx" (
    echo [ERRO] Nao encontrei Solicitacoes.xlsx nesta pasta.
    pause
    exit /b 1
)
echo [1/2] Comparando com as versoes anteriores...
node scripts\atualizar-historico.mjs
if errorlevel 1 (
    echo [ERRO] Falha ao atualizar o historico. O Node esta instalado?
    pause
    exit /b 1
)
echo.
echo [2/2] Enviando para o GitHub...
git add Solicitacoes.xlsx historico.json
git commit -m "Atualizacao dos chamados e do historico"
if errorlevel 1 (
    echo       (nada novo para enviar)
) else (
    git push
)
echo.
echo Pronto! O GitHub Pages atualiza em 1-2 minutos.
pause
