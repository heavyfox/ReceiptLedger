"""Keep Qt raster UI plugins; receipt images are decoded through Pillow.

Filter optional plugins before binary dependency discovery. The application
uses Qt Widgets raster rendering, so it needs no OpenGL fallback runtime.
"""
from pathlib import Path
from PyInstaller.utils.hooks.qt import add_qt6_dependencies

hiddenimports, binaries, datas = add_qt6_dependencies(__file__)
allowed = {"qwindows.dll", "qoffscreen.dll", "qminimal.dll", "qico.dll", "qjpeg.dll"}
binaries = [(src, dst) for src, dst in binaries if Path(src).name.lower() in allowed]
datas = [(src, dst) for src, dst in datas if Path(src).stem.endswith(("_ja", "_en"))]
