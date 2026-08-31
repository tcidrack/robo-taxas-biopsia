# -*- mode: python ; coding: utf-8 -*-
#
# Receita de build do Taxas_Biopsia.exe (lancador com auto-atualizacao).
# Este .spec e ENTRADA e VAI VERSIONADO: sem ele o build nao se reproduz noutra
# maquina (nome do .exe, collect_all, modo console).
#
#   python -m PyInstaller Taxas_Biopsia.spec --noconfirm
#
# Requisitos na maquina de build:
#   pip install pyinstaller truststore selenium pandas openpyxl requests
#
# A Analysis aponta para o LANCADOR - unico arquivo que o PyInstaller analisa.
# O issec_taxa_biopsia.py NAO entra no bundle: e baixado do GitHub em runtime.
# Por isso o lancador espelha os imports de terceiros da automacao.
#
# console=False, como nas demais automacoes da frota: o diagnostico vai para o
# logs/lancador.log ao lado do .exe (ver registrar() no lancador) e o usuario
# nao ve janela preta. Trocar para True so faz sentido ao depurar.
from PyInstaller.utils.hooks import collect_all

# selenium: traz submodulos dinamicos + o binario do Selenium Manager, que
# resolve o ChromeDriver em runtime. truststore: verificacao TLS pelo cofre de
# certificados do Windows - sem ele o download falha em rede corporativa que
# intercepta HTTPS, e a maquina congela na versao em cache.
datas = []
binaries = []
hiddenimports = []
for _pkg in ("selenium", "truststore"):
    _d, _b, _h = collect_all(_pkg)
    datas += _d
    binaries += _b
    hiddenimports += _h

# Dependencias implicitas: openpyxl e a engine que o pandas usa para .xlsx
# (nao e xlrd - esse e so para .xls legado).
hiddenimports += ["openpyxl", "pandas", "requests"]


a = Analysis(
    ["lancador_Taxas_Biopsia.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Taxas_Biopsia",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
