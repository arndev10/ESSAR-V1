import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from threading import Thread
from queue import Queue, Empty

from generador import generar


def get_plantilla_path():
    if getattr(sys, "frozen", False):
        path = os.path.join(sys._MEIPASS, "plantilla.xlsx")
    else:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla.xlsx")
    return path if os.path.isfile(path) else ""

MSG_DONE = "done"
MSG_ERROR = "error"
MSG_LOG = "log"


def run_in_thread(fn, queue):
    try:
        fn()
        queue.put((MSG_DONE, None))
    except Exception as e:
        queue.put((MSG_ERROR, str(e)))


def main():
    root = tk.Tk()
    root.title("Generador Essar")
    root.minsize(520, 420)
    root.resizable(True, True)

    plantilla_fija = get_plantilla_path()
    vars_ = {
        "insumos": tk.StringVar(value=""),
        "plantilla": tk.StringVar(value=plantilla_fija),
        "salidas": tk.StringVar(value=""),
    }

    frame = ttk.Frame(root, padding=12)
    frame.pack(fill=tk.BOTH, expand=True)

    def browse_folder(key, title):
        path = filedialog.askdirectory(title=title)
        if path:
            vars_[key].set(path)

    def browse_file():
        path = filedialog.askopenfilename(
            title="Seleccionar plantilla",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if path:
            vars_["plantilla"].set(path)

    ttk.Label(frame, text="Carpeta insumos:").grid(row=0, column=0, sticky=tk.W, pady=2)
    row0 = ttk.Frame(frame)
    row0.grid(row=1, column=0, sticky=tk.EW, pady=(0, 8))
    frame.columnconfigure(0, weight=1)
    ttk.Entry(row0, textvariable=vars_["insumos"], width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    ttk.Button(row0, text="Examinar…", command=lambda: browse_folder("insumos", "Carpeta de insumos")).pack(side=tk.RIGHT)

    ttk.Label(frame, text="Plantilla (.xlsx):").grid(row=2, column=0, sticky=tk.W, pady=2)
    row1 = ttk.Frame(frame)
    row1.grid(row=3, column=0, sticky=tk.EW, pady=(0, 8))
    entry_plantilla = ttk.Entry(row1, textvariable=vars_["plantilla"], width=50, state=tk.DISABLED if plantilla_fija else tk.NORMAL)
    entry_plantilla.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    btn_plantilla = ttk.Button(row1, text="Examinar…", command=browse_file)
    if plantilla_fija:
        btn_plantilla.configure(state=tk.DISABLED)
    btn_plantilla.pack(side=tk.RIGHT)

    ttk.Label(frame, text="Carpeta salidas:").grid(row=4, column=0, sticky=tk.W, pady=2)
    row2 = ttk.Frame(frame)
    row2.grid(row=5, column=0, sticky=tk.EW, pady=(0, 8))
    ttk.Entry(row2, textvariable=vars_["salidas"], width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    ttk.Button(row2, text="Examinar…", command=lambda: browse_folder("salidas", "Carpeta de salidas")).pack(side=tk.RIGHT)

    log_area = scrolledtext.ScrolledText(frame, height=12, state=tk.DISABLED, wrap=tk.WORD)
    log_area.grid(row=6, column=0, sticky=tk.NSEW, pady=(8, 0))
    frame.rowconfigure(6, weight=1)

    def log(msg):
        log_area.configure(state=tk.NORMAL)
        log_area.insert(tk.END, msg + "\n")
        log_area.see(tk.END)
        log_area.configure(state=tk.DISABLED)

    queue = Queue()
    btn_run = ttk.Button(frame, text="Generar")

    def do_generate():
        insumos = vars_["insumos"].get().strip()
        plantilla = vars_["plantilla"].get().strip()
        salidas = vars_["salidas"].get().strip()
        if not insumos or not plantilla or not salidas:
            log("Completa las tres rutas antes de generar.")
            return
        log_area.configure(state=tk.NORMAL)
        log_area.delete(1.0, tk.END)
        log_area.configure(state=tk.DISABLED)
        btn_run.configure(state=tk.DISABLED)

        def work():
            generar(
                insumos, plantilla, salidas,
                log_callback=lambda msg: queue.put((MSG_LOG, msg))
            )

        Thread(target=run_in_thread, args=(work, queue), daemon=True).start()

    def poll():
        try:
            while True:
                kind, payload = queue.get_nowait()
                if kind == MSG_LOG:
                    log(payload)
                elif kind == MSG_DONE:
                    log("✅ Proceso finalizado.")
                    btn_run.configure(state=tk.NORMAL)
                elif kind == MSG_ERROR:
                    log(f"❌ Error: {payload}")
                    btn_run.configure(state=tk.NORMAL)
        except Empty:
            pass
        root.after(200, poll)

    btn_run.configure(command=do_generate)
    btn_run.grid(row=7, column=0, pady=(8, 0))

    root.after(200, poll)
    root.mainloop()


if __name__ == "__main__":
    main()
