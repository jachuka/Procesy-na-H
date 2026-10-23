"""Windows entry point; the packaged executable includes Python and dependencies."""
import sys
import threading
import webbrowser
from urllib.request import urlopen
from http.server import ThreadingHTTPServer

from app import Handler


def main():
    # Let the OS choose a free port, so other local apps do not prevent startup.
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    address = f'http://127.0.0.1:{server.server_port}'
    worker = threading.Thread(target=server.serve_forever, daemon=True)
    worker.start()
    try:
        with urlopen(address, timeout=10) as response:
            if response.status != 200 or b'Excel' not in response.read():
                raise RuntimeError('Nie udało się uruchomić aplikacji.')
        if '--smoke-test' in sys.argv:
            for path in ('/app.js', '/style.css'):
                with urlopen(address + path, timeout=10) as response:
                    if response.status != 200 or not response.read():
                        raise RuntimeError('Brak plików aplikacji.')
            return
        # A native window keeps the process lifetime explicit, even after the tab closes.
        import tkinter as tk
        from tkinter import ttk
        window = tk.Tk()
        window.title('Procesy na sztukę')
        window.geometry('440x220')
        window.resizable(False, False)
        frame = ttk.Frame(window, padding=24)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Procesy na sztukę', font=('Segoe UI', 17, 'bold')).pack(anchor='w')
        ttk.Label(frame, text='Aplikacja działa w Twojej przeglądarce.\nZamknij to okno, aby ją wyłączyć.', padding=(0, 12)).pack(anchor='w')
        ttk.Button(frame, text='Otwórz aplikację', command=lambda: webbrowser.open(address)).pack(anchor='w')
        ttk.Button(frame, text='Zakończ', command=window.destroy).pack(anchor='e')
        webbrowser.open(address)
        window.mainloop()
    finally:
        server.shutdown()
        server.server_close()
        worker.join(timeout=5)


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        if '--smoke-test' not in sys.argv:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror('Procesy na sztukę', f'Nie udało się uruchomić aplikacji.\n{error}')
            root.destroy()
        sys.exit(1)
