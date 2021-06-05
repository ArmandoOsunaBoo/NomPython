import os

import base_de_datos
from modules import *
from tkinter.ttk import Progressbar
import webbrowser
from sys import exit
import datetime
root=""
git_upd="9"
import git
from git import *


class Principal:

    def our_command(self):
        pass

    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        #Configuraciones de la ventana principal
        self.master.geometry("1280x720+10+10")
        ##Primero ponemos el fondo
        self.image = Image.open("./imagenes/fondo.png")
        self.img_copy = self.image.copy()

        self.background_image = ImageTk.PhotoImage(self.image)

        self.background = tk.Label(self.frame, image=self.background_image)
        self.background.pack(fill=tk.BOTH, expand=tk.YES)
        self.background.bind('<Configure>', self.resize_image)
        #Inicio del menu a agregar en la pantalla principal
        my_menu = tk.Menu(self.master)
        self.master.config(menu=my_menu)
        ####
        menu_container = tk.Menu(my_menu)
        my_menu.add_cascade(label="Asistencias", menu=menu_container)
        menu_container.add_command(label="Genarar Reporte", command=self.wind_repAsistance)
        menu_container.add_command(label="Cargar Checadas", command=self.wind_upAsistence)
        menu_container.add_command(label="Eliminar Checadas", command=self.wind_delAsistence)
        menu_container.add_command(label="Administrar Incidencias", command=self.wind_Incidents)

        #menu_container.add_command(label="Ver Asistencias", command=self.our_command)
        menu_container.add_separator()
        menu_container.add_command(label="Exit", command=lambda: self.frame.quit())

        menu_container2 = tk.Menu(my_menu)
        my_menu.add_cascade(label="PDF", menu=menu_container2)
        menu_container2.add_command(label="Unir pdf", command=self.unir_pdf)

        menu_container3 = tk.Menu(my_menu)
        my_menu.add_cascade(label="Empleados", menu=menu_container3)
        menu_container3.add_command(label="Agregar excel empleados", command=self.agregar_info_excel)

        menu_container4 = tk.Menu(my_menu)
        my_menu.add_cascade(label="Calculo Nomina", menu=menu_container4)
        menu_container4.add_command(label="Calcular nomina", command=self.calcular_nomina)

        menu_container5 = tk.Menu(my_menu)
        my_menu.add_cascade(label="Configuraciones", menu=menu_container5)
        menu_container5.add_command(label="Actualizar Sistema", command=self.actualizar_sistema)
        menu_container5.add_command(label="Acerca de...", command=self.calcular_nomina)
        # Fin del el menú asignandolo a master
        #Fin del frame
        self.frame.pack()
        self.buscar_actualizaciones()

    def buscar_actualizaciones(self):
        pass
        bd = base_de_datos.DataBase()
        res = bd.devolver_actualizacion(git_upd)
        if res == 1:
            #parte donde se publica actualizar a el software
            self.master3 = tk.Toplevel(self.master)
            self.frame3 = tk.Frame(self.master3)
            self.master3.geometry("300x150+100+100")

            self.label_bus_act = tk.Label(self.master3, text="Hay una actualización disponible")
            self.label_bus_act.place(x=60,y=50)

            self.quitButton2 = tk.Button(self.master3, text='Actualizar', width=10,
                                        command=lambda: self.actualizar_sistema())
            self.quitButton2.place(x=40, y=100)

            self. quitButton3 = tk.Button(self.master3, text='Más tarde...', width=10,
                                        command=lambda: self.close_windows(self.master3))
            self.quitButton3.place(x=160, y=100)
            self.master3.attributes("-topmost", True)
            self.frame3.pack()


    def actualizar_sistema(self):
        pass
        print("destruccion")
        answer = messagebox.askokcancel("Question", "Seguro quiere actualizar el sistema? se cerrará el sistema")
        path = os.getcwd()
        if answer == True:
            print(os.getcwd())
            g = git.Git(os.getcwd())
            repo = git.Repo(path)
            current = repo.head.commit
            print(g.pull('origin', 'main'))
            g.pull('origin', 'main')
            if current == repo.head.commit:
                print("Repo not changed. Sleep mode activated.")
                messagebox.showwarning("Warning", "El sistema no requiere de ninguna actualización")
                return False
            else:
                messagebox.showwarning("Warning", "El sistema se ha actualizado correctamente")
                self.cierra_todo()
        else:
            pass



    def cierra_todo(self):
        pass
        root.destroy()

        ''''try:
            print(r''+path+'\'Actualizador.bat')
            subprocess.call([r''+path+"/"+'Actualizador.bat'])
        except Exception as e:
            messagebox.showwarning("Alerta", "Error al generar la actualización\n" +str(e))
        finally:
            pass
            print("salida 1")
            exit()
            print("salida 2")'''

    def calcular_nomina(self):
        pass
        filename = askopenfilename()
        rn = reporte_nomina.ReporteNomina(filename)

    def wind_delAsistence(self):
        pass
        self.master2 = tk.Toplevel(self.master)
        self.frame2 = tk.Frame(self.master2)
        self.master2.geometry("800x400+100+100")

        self.label = tk.Label(self.master2, text="Fecha Inicio")
        self.label2 = tk.Label(self.master2, text="Fecha Fin")
        self.label.place(x=175, y=50)
        self.label2.place(x=560, y=50)

        label = tk.Label(self.master2, text="Borrar registros", font=("Arial", 25))
        label.place(x=280, y=12)

        today = datetime.date.today()

        self.cal = Calendar(self.master2,
                            font="Arial 8", selectmode='day',
                            cursor="hand1", year=today.year, month=today.month, day=today.day, date_pattern='mm/dd/yyy')
        self.cal2 = Calendar(self.master2,
                             font="Arial 8", selectmode='day',
                             cursor="hand1", year=today.year, month=today.month, day=today.day,
                             date_pattern='mm/dd/yyy')

        self.cal.place(x=100, y=100)
        self.cal2.place(x=475, y=100)
        self.quitButton = tk.Button(self.master2, text='Actualizar', width=25, command=lambda: self.close_windows(self.master2))
        self.quitButton.place(x=100, y=300)

        self.quitButton = tk.Button(self.master2, text='Cancelar', width=25, command=lambda: self.close_windows(self.master2))
        self.quitButton.place(x=500, y=300)
        self.frame2.pack()

    def close_windows(self,master):
        master.destroy()

    def borrar_checadas(self):
        pass
        date1 = self.cal.get_date()
        date2 = self.cal2.get_date()

        # vamos a pasar del formato mm/dd/yyyy al formato yyyy/mm/dd
        fecha1 = date1[6:10] + "/" + date1[0:2] + "/" + date1[3:5]
        fecha2 = date2[6:10] + "/" + date2[0:2] + "/" + date2[3:5]
        print(date1)
        print(date2)
        print(fecha1)
        print(fecha2)
        bd = base_de_datos.DataBase()
        bd.borrado_asistencias(fecha1, fecha2)

    def unir_pdf(self):
        pass
        merger = PdfFileMerger()
        filez = tkinter.filedialog.askopenfilenames(parent=self.master, title='Elige los pdf a unir')
        for i in filez:
            print(i)
            input1 = open(str(i), "rb")
            merger.append(input1)
        # Write to an output PDF document
        try:
            os.mkdir("C:/tempx")
        except:
            pass
        output = open("C:/tempx/archivo.pdf", "wb")
        merger.write(output)
        path="C:/tempx/archivo.pdf"
        os.startfile(path)

    def agregar_info_excel(self):
        pass
        filename = askopenfilename()

        bd = base_de_datos.DataBase()

        bd.upload_employees(filename)


    def wind_Incidents(self):
        pass
        self.newWindow = tk.Toplevel(self.master)
        self.app = adm_Incidencias.CentroIncidencias(self.newWindow)

    def wind_repAsistance(self):
        pass
        self.newWindow = tk.Toplevel(self.master)
        self.app = mod_Nominas.ReporteAsistencia(self.newWindow)

    def wind_upAsistence(self):
        pass

        filename = askopenfilename()
        print(filename)

        bd = base_de_datos.DataBase()

        bd.upload_assistances(filename)




    def new_window(self):
        self.newWindow = tk.Toplevel(self.master)
        self.app = mod_Nominas.Nominas(self.newWindow)

    def resize_image(self,event):
        new_width = event.width
        new_height = event.height
        image = self.img_copy.resize((new_width, new_height))
        photo = ImageTk.PhotoImage(image)
        self.background.config(image=photo)
        self.background.image = photo  # avoid garbage collection



def main():
    global root
    root = tk.Tk()
    app = Principal(root)
    root.mainloop()


if __name__ == '__main__':
    main()