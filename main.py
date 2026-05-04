# =============================================================================
# main.py — Entry point da aplicação
# =============================================================================
# Fluxo:
#   1. Cria QApplication com estilo Fusion + paleta escura
#   2. Exibe LoginWindow
#   3. Quando login_aceito é emitido: fecha login, abre JanelaPortaria
# =============================================================================

import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QColor, QPalette

import ui


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Paleta escura global — idêntica ao monolito
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window,          QColor("#0f0d1a"))
    palette.setColor(QPalette.ColorRole.WindowText,      QColor("#e2e0f0"))
    palette.setColor(QPalette.ColorRole.Base,            QColor("#13101f"))
    palette.setColor(QPalette.ColorRole.AlternateBase,   QColor("#110e1e"))
    palette.setColor(QPalette.ColorRole.Text,            QColor("#e2e0f0"))
    palette.setColor(QPalette.ColorRole.Button,          QColor("#1e1b2e"))
    palette.setColor(QPalette.ColorRole.ButtonText,      QColor("#e2e0f0"))
    palette.setColor(QPalette.ColorRole.Highlight,       QColor("#7c3aed"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)

    # Mantém referência à janela principal para evitar garbage collection
    janela_principal = None

    login = ui.LoginWindow()
    login.setStyleSheet(ui.STYLESHEET)

    def abrir_sistema():
        nonlocal janela_principal
        login.close()
        janela_principal = ui.JanelaPortaria()
        janela_principal.show()

    login.login_aceito.connect(abrir_sistema)
    login.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()