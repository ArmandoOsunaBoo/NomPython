import main
from modules import *
import datetime
from tkcalendar import Calendar
import base_de_datos
from tkinter.ttk import Treeview

class CentroIncidencias:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.geometry("1024x340+100+100")

        my_menu = tk.Menu(self.master)
        self.master.config(menu=my_menu)
        ####
        menu_container = tk.Menu(my_menu)
        my_menu.add_cascade(label="Asistencias", menu=menu_container)
        menu_container.add_command(label="Cargar Incidencias", command=self.carga_incidencias)
        menu_container.add_command(label="Borrar Incidencias", command=self.borrado_incidencias)
        # menu_container.add_command(label="Ver I", command=self.our_command)
        menu_container.add_separator()
        menu_container.add_command(label="Exit", command=lambda: self.frame.quit())
        label= tk.Label(self.master,text="Ultimos registros",font=("Arial", 25))
        label.place(x=380,y=12)

        self.tv = Treeview(self.master,height=12)
        self.tv['columns'] = ('#0', '#1', '#2', '#3')
        self.tv.heading("#0", text='ID')
        self.tv.column("#0",  anchor='center', width=50)
        self.tv.heading('#1', text='No. Empleado')
        self.tv.column('#1', anchor='center', width=150)
        self.tv.heading('#2', text='Nombre')
        self.tv.column('#2', anchor='center', width=300)
        self.tv.heading('#3', text='Incidencia')
        self.tv.column('#3', anchor='center', width=150)
        self.tv.heading('#4', text='Fecha')
        self.tv.column('#4', anchor='center', width=150)
        self.tv.column('#0', stretch=tk.YES)
        self.tv.column('#1', stretch=tk.YES)
        self.tv.column('#2', stretch=tk.YES)
        self.tv.column('#3', stretch=tk.YES)
        self.tv.column('#4', stretch=tk.YES)

        self.tv.place(x=112,y=60)
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

        db = base_de_datos.DataBase()
        db.cargar_incidencias(self.tv)
        self.frame.pack()
        self.frame.after(1, lambda: self.frame.focus_force())


    def borrado_incidencias(self):
        pass
        self.master = tk.Toplevel(self.master)
        self.frame = tk.Frame(self.master)
        self.master.geometry("800x400+100+100")

        self.label = tk.Label(self.master, text="Fecha Inicio")
        self.label2 = tk.Label(self.master, text="Fecha Fin")
        self.label.place(x=175, y=50)
        self.label2.place(x=560, y=50)

        today = datetime.date.today()

        self.cal = Calendar(self.master,
                            font="Arial 8", selectmode='day',
                            cursor="hand1", year=today.year, month=today.month, day=today.day, date_pattern='mm/dd/yyy')
        self.cal2 = Calendar(self.master,
                             font="Arial 8", selectmode='day',
                             cursor="hand1", year=today.year, month=today.month, day=today.day, date_pattern='mm/dd/yyy')

        label = tk.Label(self.master, text="Borrar Incidencias", font=("Arial", 25))
        label.place(x=280, y=12)

        self.cal.place(x=100, y=100)
        self.cal2.place(x=475, y=100)
        self.quitButton = tk.Button(self.master, text='Aceptar', width=25, command= self.borrar_incidencias)
        self.quitButton.place(x=100, y=300)

        self.quitButton = tk.Button(self.master, text='Cerrar', width=25, command=self.close_windows)
        self.quitButton.place(x=500, y=300)
        self.frame.pack()

    def borrar_incidencias(self):
        pass
        date1 = self.cal.get_date()
        date2 = self.cal2.get_date()

        # vamos a pasar del formato mm/dd/yyyy al formato yyyy/mm/dd
        fecha1 = date1[6:10] + "/" + date1[0:2] + "/" + date1[3:5]
        fecha2 = date2[6:10] + "/" + date2[0:2] + "/" + date2[3:5]
        bd = base_de_datos.DataBase()
        bd.eliminado_incidencias(fecha1,fecha2,self.tv)
        self.master.after(1, lambda: self.master.focus_force())


    def close_windows(self):
        self.master.destroy()

    def carga_incidencias(self):
        pass
        filename = tk.filedialog.askopenfilename()
        print(filename)
        bd = base_de_datos.DataBase()
        bd.upload_incidences(filename,self.tv)
        self.master.after(1, lambda: self.master.focus_force())
