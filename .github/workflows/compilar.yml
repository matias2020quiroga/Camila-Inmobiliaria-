name: Compilar ejecutable para Windows

on: [workflow_dispatch]

jobs:
  build:
    runs-on: windows-latest

    steps:
    - uses: actions/checkout@v3

    - name: Configurar Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'

    - name: Instalar dependencias
      run: |
        python -m pip install --upgrade pip
        pip install pyinstaller openpyxl selenium

    - name: Crear ejecutable .EXE
      run: |
        pyinstaller --onefile --noconsole script_tandil.py

    - name: Guardar ejecutable como artefacto
      uses: actions/upload-artifact@v3
      with:
        name: Ejecutable-Tandil-Windows
        path: dist/script_tandil.exe
