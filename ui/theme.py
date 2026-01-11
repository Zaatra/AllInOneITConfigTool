"""Dark theme styling helper."""
from __future__ import annotations

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QStyleFactory


def apply_dark_theme() -> None:
    app = QApplication.instance()
    if app is None:
        return

    app.setStyle(QStyleFactory.create("Fusion"))

    background = QColor("#1e1e1e")
    surface = QColor(32, 32, 32)
    text = QColor(230, 230, 230)
    accent = QColor("#007acc")
    disabled_text = QColor(130, 130, 130)

    palette = QPalette()
    palette.setColor(QPalette.Window, background)
    palette.setColor(QPalette.WindowText, text)
    palette.setColor(QPalette.Base, QColor(18, 18, 18))
    palette.setColor(QPalette.AlternateBase, surface)
    palette.setColor(QPalette.ToolTipBase, surface)
    palette.setColor(QPalette.ToolTipText, text)
    palette.setColor(QPalette.Text, text)
    palette.setColor(QPalette.Button, surface)
    palette.setColor(QPalette.ButtonText, text)
    palette.setColor(QPalette.BrightText, QColor("#ff4081"))
    palette.setColor(QPalette.Highlight, accent)
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    palette.setColor(QPalette.Disabled, QPalette.Text, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.ButtonText, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.WindowText, disabled_text)
    app.setPalette(palette)
