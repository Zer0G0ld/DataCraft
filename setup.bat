@echo off
echo ========================================
echo DataCraft - Configuracao do Ambiente
echo ========================================
echo.

REM Verifica se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado!
    echo Por favor, instale o Python em: https://www.python.org/downloads/
    echo Nao esqueca de marcar "Add Python to PATH" durante a instalacao!
    pause
    exit /b 1
)

echo Python encontrado!
python --version
echo.

REM Cria ambiente virtual se não existir
if not exist "venv" (
    echo Criando ambiente virtual...
    python -m venv venv
    echo Ambiente virtual criado!
) else (
    echo Ambiente virtual ja existe!
)
echo.

REM Ativa o ambiente virtual
echo Ativando ambiente virtual...
call venv\Scripts\activate

REM Verifica se pip está funcionando
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Pip nao encontrado!
    pause
    exit /b 1
)

echo Pip encontrado!
echo.

REM Instala dependências
echo Instalando dependencias...
python -m pip install --upgrade pip
python -m pip install pandas==2.2.0
python -m pip install numpy==1.26.3
python -m pip install openpyxl==3.1.2
python -m pip install PyYAML==6.0.1
python -m pip install sqlalchemy==2.0.25
python -m pip install pyarrow==15.0.0
python -m pip install xlrd==2.0.1
python -m pip install markdown==3.5.2
python -m pip install jinja2==3.1.3
python -m pip install pyinstaller==6.4.0

echo.
echo ========================================
echo Configuracao concluida com sucesso!
echo Para executar o DataCraft:
echo 1. Execute: venv\Scripts\activate
echo 2. Execute: python DataCraft.py
echo ========================================
pause