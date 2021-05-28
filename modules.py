#Modulo de tkinter para ventanas
import tkinter as tk
import mod_Nominas
import mod_Utilidades
import subprocess
from PIL import ImageTk, Image
import interfaces as interfaces
from reporte_Asistencias import *
import base_de_datos
import adm_Incidencias
from tkcalendar import Calendar
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
import utilidades_Generales
import PyPDF3
import tkinter.filedialog
from PyPDF3 import PdfFileMerger
import reporte_nomina
from tkinter import messagebox
from git import Repo