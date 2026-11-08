import os
import sys
import time
import threading
import webbrowser

import streamlit.web.cli as stcli


def obtener_ruta_recursos():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def obtener_ruta_ejecutable():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def abrir_navegador():
    time.sleep(3)
    webbrowser.open("http://localhost:8501")


if __name__ == "__main__":
    recursos_dir = obtener_ruta_recursos()
    exe_dir = obtener_ruta_ejecutable()

    app_path = os.path.join(recursos_dir, "app.py")

    os.environ["AUTOINSPECCIONES_DATOS"] = os.path.join(exe_dir, "datos")

    threading.Thread(target=abrir_navegador, daemon=True).start()

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.address=localhost",
        "--server.port=8501"
    ]

    sys.exit(stcli.main())