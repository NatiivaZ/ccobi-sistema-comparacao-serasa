@echo off
cd /d "%~dp0"
chcp 65001 >nul
echo ========================================
echo   Sistema de Analise SERASA x Divida Ativa
echo   Autos de Infracao ANTT
echo ========================================
echo.
echo Verificando Python...
python --version
if errorlevel 1 (
    echo ERRO: Python nao encontrado! Instale o Python primeiro.
    pause
    exit /b 1
)
echo.
echo Instalando/Atualizando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERRO: Falha ao instalar dependencias!
    pause
    exit /b 1
)
echo.
echo Configurando Streamlit...
python "%~dp0..\config_streamlit.py"
echo.
echo ========================================
echo Iniciando sistema...
echo ========================================
echo.
echo O navegador abrira automaticamente.
echo Para parar o sistema, pressione Ctrl+C
echo.
python -m streamlit run app.py --server.headless false
pause
