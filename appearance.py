"""Application themes and large, explicit receipt selection indicators."""

from PySide6.QtCore import QEvent, QRect, QSize, Qt
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPalette, QPen
from PySide6.QtWidgets import QStyledItemDelegate, QStyle, QStyleOptionViewItem


def system_theme():
    """Read the Windows app color preference without modifying system settings."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "dark" if value == 0 else "light"
    except (ImportError, OSError):
        return "light"


COLORS = {
    "light": dict(bg="#f4f7fa", surface="#ffffff", text="#17324a", muted="#486277",
                  border="#9cabb9", button="#e5edf3", hover="#d3e7f2", accent="#17698a",
                  accent_text="#ffffff", selected="#d4eaf5", disabled="#e8edf2", dim="#6d7d8c",
                  warning="#885100", preview="#e4ebf1", check="#08765b", check_text="#ffffff"),
    "dark": dict(bg="#161d27", surface="#202b39", text="#edf3fa", muted="#b2c2d3",
                 border="#687c92", button="#2c3b4e", hover="#3a5068", accent="#318db5",
                 accent_text="#ffffff", selected="#304f6c", disabled="#25303f", dim="#97a7b9",
                 warning="#ffd18a", preview="#111822", check="#49d6b1", check_text="#062b24"),
}


def apply_theme(window, theme):
    c = COLORS.get(theme, COLORS["light"])
    palette = QPalette()
    for role, key in ((QPalette.ColorRole.Window, "bg"), (QPalette.ColorRole.WindowText, "text"),
                      (QPalette.ColorRole.Base, "surface"), (QPalette.ColorRole.AlternateBase, "bg"),
                      (QPalette.ColorRole.Text, "text"), (QPalette.ColorRole.Button, "button"),
                      (QPalette.ColorRole.ButtonText, "text"), (QPalette.ColorRole.Highlight, "selected"),
                      (QPalette.ColorRole.HighlightedText, "text"), (QPalette.ColorRole.ToolTipBase, "surface"),
                      (QPalette.ColorRole.ToolTipText, "text"), (QPalette.ColorRole.PlaceholderText, "muted"),
                      (QPalette.ColorRole.Link, "accent"), (QPalette.ColorRole.Mid, "border"),
                      (QPalette.ColorRole.Dark, "border"), (QPalette.ColorRole.Light, "surface")):
        palette.setColor(role, QColor(c[key]))
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(c["dim"]))
    window.setPalette(palette)
    window.setStyleSheet("""
        QMainWindow, QWidget { background: %(bg)s; color: %(text)s; font-size: 13px; }
        QGroupBox { border: 1px solid %(border)s; border-radius: 7px; margin-top: 12px; padding: 12px; font-weight: bold; }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
        QLineEdit, QComboBox, QSpinBox, QTableWidget, QListWidget, QScrollArea {
            background: %(surface)s; color: %(text)s; border: 1px solid %(border)s;
            border-radius: 4px; padding: 4px; selection-background-color: %(selected)s; selection-color: %(text)s;
        }
        QLineEdit:focus, QComboBox:focus, QTableWidget:focus, QListWidget:focus { border: 2px solid %(accent)s; }
        QLineEdit:disabled, QComboBox:disabled { background: %(disabled)s; color: %(dim)s; }
        QComboBox QAbstractItemView { background: %(surface)s; color: %(text)s; selection-background-color: %(selected)s; selection-color: %(text)s; }
        QHeaderView::section, QTableCornerButton::section { background: %(button)s; color: %(text)s; border: 1px solid %(border)s; padding: 5px; }
        QTableWidget { gridline-color: %(border)s; }
        QPushButton { background: %(button)s; color: %(text)s; border: 1px solid %(border)s; border-radius: 5px; padding: 7px 12px; }
        QPushButton:hover { background: %(hover)s; }
        QPushButton:focus { border: 2px solid %(accent)s; }
        QPushButton:disabled { background: %(disabled)s; color: %(dim)s; border-color: %(border)s; }
        QPushButton#primary { background: %(accent)s; color: %(accent_text)s; border-color: %(accent)s; font-weight: bold; }
        QPushButton#primary:hover { background: %(hover)s; color: %(text)s; }
        QPushButton#primary:disabled { background: %(disabled)s; color: %(dim)s; border-color: %(border)s; }
        QTabWidget::pane { border: 0; }
        QTabBar::tab { padding: 10px 18px; background: %(bg)s; color: %(muted)s; }
        QTabBar::tab:selected { background: %(surface)s; color: %(text)s; font-weight: bold; border-bottom: 3px solid %(accent)s; }
        QTabBar::tab:disabled { color: %(dim)s; }
        QProgressBar { background: %(surface)s; color: %(text)s; border: 1px solid %(border)s; border-radius: 4px; text-align: center; }
        QProgressBar::chunk { background: %(selected)s; }
        QScrollBar:vertical { background: %(bg)s; width: 14px; margin: 0; }
        QScrollBar:horizontal { background: %(bg)s; height: 14px; margin: 0; }
        QScrollBar::handle { background: %(border)s; border-radius: 5px; min-height: 24px; min-width: 24px; }
        QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
        QScrollBar::add-page, QScrollBar::sub-page { background: %(bg)s; }
        QToolTip { background: %(surface)s; color: %(text)s; border: 1px solid %(border)s; padding: 6px; }
        QLabel#receiptPreview { background: %(preview)s; color: %(muted)s; }
        QLabel#receiptWarning { color: %(warning)s; padding: 6px; }
        QLabel#categorySummary { color: %(muted)s; padding: 0 8px 8px 8px; }
    """ % c)


class ReceiptCheckDelegate(QStyledItemDelegate):
    """Paint and hit-test a 26px checkbox, independent of Windows theme defaults."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.theme = "light"

    def indicator_rect(self, option):
        return QRect(option.rect.left() + 8, option.rect.center().y() - 13, 26, 26)

    def sizeHint(self, option, index):
        size = super().sizeHint(option, index)
        return QSize(size.width() + 26, max(size.height(), 44))

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        c = COLORS[self.theme]
        selected = bool(opt.state & QStyle.StateFlag.State_Selected)
        available = bool(index.flags() & Qt.ItemFlag.ItemIsUserCheckable)
        checked = index.data(Qt.ItemDataRole.CheckStateRole) == Qt.CheckState.Checked.value
        rect = self.indicator_rect(opt)
        painter.save()
        painter.fillRect(opt.rect, QColor(c["selected"] if selected else c["surface"]))
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(c["check"] if checked else c["border"]), 2.5))
        painter.setBrush(QColor(c["check"] if checked else c["surface"] if available else c["disabled"]))
        painter.drawRoundedRect(rect.adjusted(1, 1, -1, -1), 4, 4)
        if checked:
            pen = QPen(QColor(c["check_text"]), 3.5, Qt.PenStyle.SolidLine,
                       Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            tick = QPainterPath()
            tick.moveTo(rect.left() + 6, rect.top() + 13)
            tick.lineTo(rect.left() + 11, rect.top() + 18)
            tick.lineTo(rect.left() + 21, rect.top() + 7)
            painter.drawPath(tick)
        elif not available:
            painter.setPen(QPen(QColor(c["dim"]), 2))
            painter.drawLine(rect.left() + 8, rect.center().y(), rect.right() - 8, rect.center().y())
        text_rect = opt.rect.adjusted(46, 0, -8, 0)
        painter.setFont(opt.font)
        painter.setPen(QColor(c["text"]))
        text = opt.fontMetrics.elidedText(opt.text, Qt.TextElideMode.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, text)
        if opt.state & QStyle.StateFlag.State_HasFocus:
            painter.setPen(QPen(QColor(c["accent"]), 1, Qt.PenStyle.DotLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(opt.rect.adjusted(1, 1, -2, -2))
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if not (index.flags() & Qt.ItemFlag.ItemIsUserCheckable and index.flags() & Qt.ItemFlag.ItemIsEnabled):
            return False
        mouse_events = (QEvent.Type.MouseButtonPress, QEvent.Type.MouseButtonRelease, QEvent.Type.MouseButtonDblClick)
        if event.type() in mouse_events:
            if event.button() != Qt.MouseButton.LeftButton or not self.indicator_rect(option).contains(event.position().toPoint()):
                return False
            if event.type() != QEvent.Type.MouseButtonRelease:
                return True
        elif event.type() == QEvent.Type.KeyPress:
            if event.key() not in (Qt.Key.Key_Space, Qt.Key.Key_Select):
                return False
        else:
            return False
        checked = index.data(Qt.ItemDataRole.CheckStateRole) == Qt.CheckState.Checked.value
        return model.setData(index, Qt.CheckState.Unchecked.value if checked else Qt.CheckState.Checked.value,
                             Qt.ItemDataRole.CheckStateRole)
