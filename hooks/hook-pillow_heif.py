from PyInstaller.utils.hooks import collect_dynamic_libs

binaries = collect_dynamic_libs("pillow_heif")
hiddenimports = ["_pillow_heif"]
