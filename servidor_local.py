from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import threading
import webbrowser
import os

PORT = 8000
BASE_DIR = Path(__file__).resolve().parent
os.chdir(BASE_DIR)


def open_browser():
    webbrowser.open(f'http://localhost:{PORT}/login.html')


if __name__ == '__main__':
    print(f'Servidor local iniciado em http://localhost:{PORT}/login.html')
    threading.Timer(1.0, open_browser).start()
    server = ThreadingHTTPServer(('127.0.0.1', PORT), SimpleHTTPRequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServidor encerrado.')
    finally:
        server.server_close()
