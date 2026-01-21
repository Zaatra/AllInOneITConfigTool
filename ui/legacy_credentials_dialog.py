"""Credentials prompt for accessing legacy driver repositories."""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QVBoxLayout,
    QWidget,
)


class LegacyCredentialsDialog(QDialog):
    def __init__(self, share: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._share = share
        self._username = QLineEdit()
        self._password = QLineEdit()
        self.setWindowTitle("Legacy Repo Credentials")
        self.setMinimumWidth(420)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        info = QLabel(f"Credentials are required to access:\n{self._share}")
        info.setWordWrap(True)
        layout.addWidget(info)

        form = QFormLayout()
        self._username.setPlaceholderText("DOMAIN\\username or username")
        self._password.setEchoMode(QLineEdit.Password)
        form.addRow("Username", self._username)
        form.addRow("Password", self._password)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def credentials(self) -> tuple[str, str]:
        return (self._username.text().strip(), self._password.text())
