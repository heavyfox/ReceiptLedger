"""ReceiptLedger uses QWidget/QPainter raster rendering, never OpenGL/QML.

Core DLL dependencies are collected by each Qt module hook. Do not collect the
optional 20 MiB Mesa OpenGL runtime or the Direct3D shader compiler. Keep the
binding helpers required by PySide6's Python operator implementations.
"""
from PyInstaller.utils.hooks.qt import ensure_single_qt_bindings_package

ensure_single_qt_bindings_package("PySide6")
hiddenimports = ["shiboken6", "inspect", "PySide6.support.deprecated"]
