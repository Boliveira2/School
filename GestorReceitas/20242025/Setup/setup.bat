@echo off
SET PYTHON_VERSION=3.10.8
SET PYTHON_INSTALLER=python-%PYTHON_VERSION%-amd64.exe
SET POWERSHELL_VERSION=7.3.0
SET POWERSHELL_INSTALLER=PowerShell-%POWERSHELL_VERSION%-win-x64.msi

:: Verificar se o Python está instalado
python --version >nul 2>&1
IF NOT %ERRORLEVEL% == 0 (
    echo Python não encontrado, instalando o Python %PYTHON_VERSION%...

    :: Baixar o instalador do Python
    curl -O https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_INSTALLER%

    :: Instalar o Python silenciosamente (modo quiet)
    %PYTHON_INSTALLER% /quiet InstallAllUsers=1 PrependPath=1

    :: Remover o instalador baixado
    del %PYTHON_INSTALLER%

    echo Python instalado com sucesso.
) ELSE (
    echo Python já está instalado.
)

:: Atualizar o pip para garantir a versão mais recente
python -m ensurepip --upgrade
python -m pip install --upgrade pip

:: Verificar se o arquivo requirements.txt existe
IF NOT EXIST requirements.txt (
    echo O arquivo requirements.txt não foi encontrado. Certifique-se de que ele está no mesmo diretório.
    exit /b 1
)

:: Instalar as dependências listadas em requirements.txt
pip install -r requirements.txt

:: Verificar se o PowerShell 7 está instalado
powershell -Command "Get-Command pwsh -ErrorAction SilentlyContinue"
IF NOT %ERRORLEVEL% == 0 (
    echo PowerShell 7 não encontrado, instalando PowerShell 7...

    :: Baixar o instalador do PowerShell 7
    curl -O https://github.com/PowerShell/PowerShell/releases/download/v%POWERSHELL_VERSION%/%POWERSHELL_INSTALLER%

    :: Instalar o PowerShell 7 silenciosamente
    msiexec /i %POWERSHELL_INSTALLER% /quiet /norestart

    :: Remover o instalador baixado
    del %POWERSHELL_INSTALLER%

    echo PowerShell 7 instalado com sucesso.
) ELSE (
    echo PowerShell 7 já está instalado.
)

:: Mensagem final
echo Instalação concluída com sucesso!
pause
