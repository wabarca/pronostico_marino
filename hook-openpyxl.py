from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Incluir todos los submódulos de openpyxl
hiddenimports = collect_submodules('openpyxl')

# Incluir todos los archivos de datos de openpyxl
datas = collect_data_files('openpyxl')
