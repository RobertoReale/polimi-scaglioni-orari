#!/usr/bin/env python3
"""
Avvio del programma (Windows, Linux, macOS).
Controlla che ci sia tutto il necessario, se manca qualche libreria la installa
in una cartella locale (.venv) e apre l'interfaccia grafica.

    python avvia.py              interfaccia grafica
    python avvia.py --terminale  mostra i comandi per l'uso da terminale
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
VENV = HERE / ".venv"
LIBRERIE = {"requests": "requests", "bs4": "beautifulsoup4", "openpyxl": "openpyxl"}


def mancanti():
    out = []
    for modulo, pacchetto in LIBRERIE.items():
        try:
            __import__(modulo)
        except ImportError:
            out.append(pacchetto)
    return out


def python_venv():
    return VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def installa():
    """Crea .venv e ci installa le librerie; poi riavvia questo script con quel Python."""
    print("Prima esecuzione: installo le librerie necessarie (serve internet, una volta sola)…")
    py = python_venv()
    if not py.exists() or subprocess.run([str(py), "-m", "pip", "--version"], capture_output=True).returncode:
        shutil.rmtree(VENV, ignore_errors=True)
        r = subprocess.run([sys.executable, "-m", "venv", str(VENV)], capture_output=True, text=True)
        if r.returncode != 0 or not py.exists():
            shutil.rmtree(VENV, ignore_errors=True)  # non lasciare un ambiente a metà
            print("\nNon riesco a creare l'ambiente Python locale (modulo venv mancante).")
            if sys.platform.startswith("linux"):
                print("Su Ubuntu/Debian installa i pacchetti necessari con UNO di questi comandi, poi riavvia:\n"
                      "    sudo apt install python3-venv python3-tk\n"
                      "oppure (librerie di sistema, senza ambiente locale):\n"
                      "    sudo apt install python3-requests python3-bs4 python3-openpyxl python3-tk\n"
                      "Fedora:  sudo dnf install python3-tkinter      Arch:  sudo pacman -S tk")
            sys.exit(1)
    r = subprocess.run([str(py), "-m", "pip", "install", "--disable-pip-version-check", "-q",
                        "-r", str(HERE / "requirements.txt")])
    if r.returncode != 0:
        print("\nInstallazione delle librerie non riuscita: controlla la connessione e riprova.")
        sys.exit(1)
    riavvia_con(py)


def riavvia_con(py):
    sys.exit(subprocess.run([str(py), str(HERE / "avvia.py"), *sys.argv[1:]]).returncode)


def main():
    if sys.version_info < (3, 8):
        sys.exit("Serve Python 3.8 o più recente: scaricalo da https://www.python.org/downloads/")
    os.chdir(HERE)
    sys.path.insert(0, str(HERE))

    if mancanti():
        # se esiste già l'ambiente locale, usa quello
        if python_venv().exists() and Path(sys.prefix).resolve() != VENV.resolve():
            riavvia_con(python_venv())
        installa()

    if "--terminale" in sys.argv:
        print(__doc__)
        print("Scaricare i dati:   python scarica_manifesti.py --help   (--elenca mostra i corsi disponibili)")
        print("Esportare i dati:   python esporta.py --help")
        return

    try:
        import tkinter  # noqa: F401
    except ImportError:
        print("Manca il modulo grafico di Python (tkinter).")
        if sys.platform.startswith("linux"):
            print("Installalo con uno di questi comandi, poi riavvia:\n"
                  "    Ubuntu/Debian:  sudo apt install python3-tk\n"
                  "    Fedora:         sudo dnf install python3-tkinter\n"
                  "    Arch:           sudo pacman -S tk")
        print("\nIn alternativa puoi usare il programma da terminale:  python avvia.py --terminale")
        sys.exit(1)

    import interfaccia
    interfaccia.main()


if __name__ == "__main__":
    main()
