#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
from BIN.Fixer import Arreglar_Comprobantes
import os

def Abrir_Excel():
    os.system('start EXCEL.EXE "Lista de Archivos.xlsx"')


def donaciones():
    os.system("start https://cafecito.app/abustos")



class GuiEnvíoDeMailsApp:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=250,
            width=325)
        Toplevel_1.iconbitmap("BIN/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.overrideredirect("False")
        Toplevel_1.resizable(False, False)
        Toplevel_1.title("Arreglo de archivos")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="BIN/ABP blanco sin fondo.png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            foreground="#ffffff",
            justify="center",
            takefocus=False,
            text='Arreglo de "Importes Totales" y "Fechas" de archivos de compras y ventas para importar al Nacional\n',
            wraplength=325)
        Label_1.pack(expand=True, side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand=True, side="top")
        self.Excel = ttk.Button(Toplevel_1)
        self.Excel.configure(text='Abrir Excel con datos' , command = Abrir_Excel)
        self.Excel.pack(expand=True, padx=0, pady=4, side="top")
        self.Enviar = ttk.Button(Toplevel_1)
        self.Enviar.configure(text='Corregir' , command=Arreglar_Comprobantes)
        self.Enviar.pack(expand=True, pady=4, side="top")
        self.Colaboraciones = ttk.Button(Toplevel_1)
        self.Colaboraciones.configure(text='Colaboraciones' , command=donaciones)
        self.Colaboraciones.pack(expand=True, pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = GuiEnvíoDeMailsApp()
    app.run()
