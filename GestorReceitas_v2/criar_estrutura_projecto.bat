@echo off
SET PROJETO=GestorReceitas_v2

echo.
echo 🔧 A criar estrutura do projeto %PROJETO%...
mkdir %PROJETO%
cd %PROJETO%

:: Criar subpastas
mkdir src
mkdir tests
mkdir .vscode

:: Criar estrutura src/gestor e subpastas
mkdir src\gestor
mkdir src\gestor\core
mkdir src\gestor\relatorio
mkdir src\gestor\comunicacao
mkdir src\gestor\cli

:: Criar ambiente virtual
echo 🔄 A criar ambiente virtual...
python -m venv venv

:: Criar ficheiros base

echo 📄 A criar ficheiros iniciais...

:: ficheiros __init__.py vazios para pacotes
type nul > src\gestor\__init__.py
type nul > src\gestor\core\__init__.py
type nul > src\gestor\relatorio\__init__.py
type nul > src\gestor\comunicacao\__init__.py
type nul > src\gestor\cli\__init__.py
type nul > tests\test_basico.py

:: Criar config.py com exemplos mínimos
echo # Configurações base e caminhos > src\gestor\config.py
echo BASE_DIR = "%cd%" >> src\gestor\config.py
echo MESES = [^"janeiro^", ^"fevereiro^", ^"março^", ^"abril^", ^"maio^", ^"junho^", ^"julho^", ^"agosto^", ^"setembro^", ^"outubro^", ^"novembro^", ^"dezembro^"] >> src\gestor\config.py

:: Criar main simples em cli\gerar_relatorio.py
(
echo def main():
echo.    print("Este é o script para gerar relatório mensal")
echo.
echo if __name__ == "__main__":
echo.    main()
) > src\gestor\cli\gerar_relatorio.py

:: Criar main simples em cli\gerar_transferencias.py
(
echo def main():
echo.    print("Este é o script para gerar transferências")
echo.
echo if __name__ == "__main__":
echo.    main()
) > src\gestor\cli\gerar_transferencias.py

:: Criar main simples em cli\gerar_emails.py
(
echo def main():
echo.    print("Este é o script para gerar emails")
echo.
echo if __name__ == "__main__":
echo.    main()
) > src\gestor\cli\gerar_emails.py

:: Criar requirements.txt
echo pandas> requirements.txt
echo openpyxl>> requirements.txt
echo odfpy>> requirements.txt

:: settings.json do VS Code
echo {> .vscode\settings.json
echo.  "python.pythonPath": "venv\\Scripts\\python.exe",>> .vscode\settings.json
echo.  "python.formatting.provider": "black",>> .vscode\settings.json
echo.  "python.linting.enabled": true,>> .vscode\settings.json
echo.  "python.linting.flake8Enabled": true,>> .vscode\settings.json
echo.  "editor.formatOnSave": true,>> .vscode\settings.json
echo.  "python.testing.unittestEnabled": true,>> .vscode\settings.json
echo.  "python.testing.autoTestDiscoverOnSaveEnabled": true>> .vscode\settings.json
echo }>> .vscode\settings.json

:: Script run.bat para executar a aplicação principal (exemplo gerar_relatorio)
echo @echo off> run.bat
echo call venv\Scripts\activate>> run.bat
echo python src\gestor\cli\gerar_relatorio.py>> run.bat
echo pause>> run.bat

echo.
echo 📦 A instalar dependências base...
call venv\Scripts\activate
pip install -r requirements.txt

echo.
echo ✅ Projeto %PROJETO% criado com sucesso!
pause
