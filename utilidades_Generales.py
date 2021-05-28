import modules
import tkinter as tk
import main
class ventana_carga:
    def __init__(self):
        pass
        self.master = tk.Toplevel(main.root)
        self.frame2 = tk.Frame(self.master)
        self.master.geometry("400x200+100+100")
        self.frame2.pack()
        self.frame2.grab_set()
        label = tk.Label(self.master, text="Cargando datos porfavor espere...")
        label.place(x=100, y=50)
        self.my_progress = tk.ttk.Progressbar(self.master, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.my_progress.place(x=50, y=100)
        self.my_progress.update_idletasks()
        self.my_progress.update()
        self.frame2.after(1, lambda: self.frame2.focus_force())
        self.frame2.update_idletasks()

    def cargar_valor(self, valor):
        pass
        self.my_progress['value'] = valor
        self.my_progress.update_idletasks()
        self.my_progress.update()

    def cerrar_ventana(self):
        pass
        self.master.grab_release()
        self.master.destroy()

