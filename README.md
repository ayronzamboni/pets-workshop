# Editor de Planilha Simples (Windows)

Resumo
Este é um pequeno aplicativo desktop em Python (PySimpleGUI) que permite abrir um arquivo Excel (.xlsx) ou CSV, visualizar e editar linhas, adicionar/remover linhas, filtrar por texto e salvar de volta (com backup automático).

Pré-requisitos
- Windows
- Python 3.10+ (somente se você quiser rodar sem o .exe)
- pip

Instalação rápida
1. Crie uma pasta.
2. Salve os arquivos `main.py`, `requirements.txt` e `build_exe.bat`.
3. Abra o terminal nessa pasta.
4. Rode:
   pip install -r requirements.txt

Executar
python main.py

Gerar executável
pip install pyinstaller
build_exe.bat

O executável será gerado em:
dist\EditorPlanilha.exe

Funcionalidades
- Abrir arquivo .xlsx/.csv
- Navegar por abas
- Editar linha
- Adicionar linha
- Excluir linha
- Filtrar por texto
- Salvar e salvar como
- Backup automático antes de sobrescrever
