import tkinter

from modules import *

class Asistencias:

    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        #Creamos el menú asignandolo a master
        my_menu = tk.Menu(self.master)
        self.master.config(menu=my_menu)
        ####
        menu_container = tk.Menu(my_menu)
        my_menu.add_cascade(label="Inicio...", menu=menu_container)
        menu_container.add_command(label="Nuevo..", command=self.our_command)
        menu_container.add_separator()
        menu_container.add_command(label="Exit", command=lambda:self.frame.quit())

        menu_container2 = tkinter.Menu(my_menu)
        my_menu.add_cascade(label="Editar",menu=menu_container2)
        menu_container2.add_command(label="Unir pdf", command=self.our_command)
        menu_container2.add_command(label="Imprimir Pdf", command=self.our_command)
        # Fin del el menú asignandolo a master

        self.quitButton = tk.Button(self.frame, text = 'Quit', width = 25, command = self.close_windows)
        self.quitButton.pack()
        self.frame.pack()

    def our_command(self):
        pass
        print("hello world")

    def close_windows(self):
        self.master.destroy()