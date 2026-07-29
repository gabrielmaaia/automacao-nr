"""
main.py — Ponto de entrada do aplicativo Matriz de Treinamentos NR.

Inicia a interface gráfica CustomTkinter.
Para compilar em .exe, aponte o PyInstaller para este arquivo.
"""

from gui import MatrizApp


def main():
    app = MatrizApp()
    app.mainloop()


if __name__ == "__main__":
    main()
