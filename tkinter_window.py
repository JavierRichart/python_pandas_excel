# drop_xlsx_ui.py
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from tkinterdnd2 import TkinterDnD, DND_FILES


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Arrastra tu Excel (.xlsx)")
        self.geometry("560x300")
        self.resizable(False, False)

        self.selected_path = tk.StringVar(value="Ningún archivo seleccionado")

        wrap = ttk.Frame(self, padding=12)
        wrap.pack(fill="both", expand=True)

        ttk.Label(wrap, text="Arrastra un archivo .xlsx a la zona inferior").pack(anchor="w")

        self.drop_zone = tk.Frame(wrap, width=520, height=160, bg="#fafafa", highlightthickness=2,
                                  highlightbackground="#9aa4af")
        self.drop_zone.pack(pady=10, fill="x")
        self.drop_zone.pack_propagate(False)

        self.hint = ttk.Label(self.drop_zone, text="Suelta aquí tu .xlsx", foreground="#6b7280")
        self.hint.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<Drop>>", self.on_drop)

        ttk.Label(wrap, textvariable=self.selected_path, foreground="#374151").pack(anchor="w", pady=(8, 0))

        self.btn_procesar = ttk.Button(wrap, text="Procesar", command=self.on_process, state="disabled")
        self.btn_procesar.pack(anchor="e", pady=(10, 0))

    def on_drop(self, event):
        try:
            files = self.tk.splitlist(event.data)
            if not files:
                return
            p = Path(files[0])
            if p.suffix.lower() != ".xlsx":
                messagebox.showwarning("Formato no válido", "Por favor, suelta un archivo con extensión .xlsx")
                return
            self.selected_path.set(str(p))
            self.btn_procesar.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_process(self):
        messagebox.showinfo("Pendiente",
                            f"Archivo seleccionado:\n{self.selected_path.get()}\n\nLa lógica de proceso se añadirá más adelante.")


if __name__ == "__main__":
    App().mainloop()
