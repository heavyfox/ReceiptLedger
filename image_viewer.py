"""Receipt inspection with zoom, pan, fit and non-destructive rotation."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QPixmap, QTransform
from PySide6.QtWidgets import QGraphicsView, QGraphicsScene, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget


class ReceiptView(QGraphicsView):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.owner.fitted:
            self.owner.fit()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.owner.zoom(1.25 if event.angleDelta().y() > 0 else 0.8)
            event.accept()
        else:
            super().wheelEvent(event)


class ImageViewer(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source = QPixmap()
        self.angle = 0
        self.fitted = True
        self.item = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        bar = QHBoxLayout()
        self.controls = []
        for text, action, tip in (
            ("−", lambda: self.zoom(0.8), "縮小"),
            ("＋", lambda: self.zoom(1.25), "拡大（Ctrl＋ホイールでも操作できます）"),
            ("全体表示", self.fit, "画像全体を表示領域に合わせる"),
            ("100%", self.actual_size, "プレビュー画像を等倍で表示"),
            ("左回転", lambda: self.rotate(-90), "表示を左に90度回転。元画像は変更しません"),
            ("右回転", lambda: self.rotate(90), "表示を右に90度回転。元画像は変更しません"),
        ):
            button = QPushButton(text)
            button.setToolTip(tip)
            button.clicked.connect(action)
            bar.addWidget(button)
            self.controls.append(button)
        self.scale_label = QLabel()
        bar.addWidget(self.scale_label)
        layout.addLayout(bar)
        self.view = ReceiptView(self)
        self.view.setMinimumSize(240, 180)
        layout.addWidget(self.view, 1)
        self.message = QLabel("画像を追加してください")
        self.message.setWordWrap(True)
        layout.addWidget(self.message)
        self._enabled(False)

    def _enabled(self, enabled):
        for button in self.controls:
            button.setEnabled(enabled)

    def pixmap(self):
        return self.source

    def setPixmap(self, pixmap):
        self.source = pixmap
        self.angle = 0
        self.view.scene().clear()
        self.item = None
        if not pixmap.isNull():
            self.item = self.view.scene().addPixmap(pixmap)
            self.view.scene().setSceneRect(self.item.boundingRect())
            self.message.setText("ドラッグで移動・Ctrl＋ホイールで拡大／縮小")
        self._enabled(not pixmap.isNull())
        self.fit()

    def setText(self, text):
        self.message.setText(text)

    def _scale_text(self):
        self.scale_label.setText(f"{self.view.transform().m11() * 100:.0f}%" if self.item else "")

    def fit(self):
        self.fitted = True
        if self.item:
            self.view.fitInView(self.item, Qt.AspectRatioMode.KeepAspectRatio)
        self._scale_text()

    def actual_size(self):
        self.fitted = False
        self.view.resetTransform()
        self._scale_text()

    def zoom(self, multiplier):
        if not self.item:
            return
        self.fitted = False
        current = self.view.transform().m11()
        target = min(4.0, max(0.05, current * multiplier))
        self.view.scale(target / current, target / current)
        self._scale_text()

    def rotate(self, degrees):
        if not self.item:
            return
        self.angle = (self.angle + degrees) % 360
        self.item.setPixmap(self.source.transformed(QTransform().rotate(self.angle), Qt.TransformationMode.SmoothTransformation))
        self.view.scene().setSceneRect(self.item.boundingRect())
        if self.fitted:
            self.fit()

    def state(self):
        return {"angle": self.angle, "fitted": self.fitted, "scale": self.view.transform().m11()}

    def restore_state(self, state):
        self.rotate(int(state.get("angle", 0)) % 360)
        if not state.get("fitted", True):
            self.actual_size()
            self.zoom(float(state.get("scale", 1)))
