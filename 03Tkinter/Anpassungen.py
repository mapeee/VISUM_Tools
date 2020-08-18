# -*- coding: utf-8 -*-
"""
Created on Thu Jul 30 09:24:56 2020

@author: mape
"""
import win32com.client.dynamic
import tkinter as tk
from tkinter import filedialog as fd
from pathlib import Path
from tkinter import ttk

class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):       
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        
        ##connection to path
        self.f = open(Path.home() / 'python32' / 'python_dir.txt', mode='r')
        self.path = Path.joinpath(Path(self.f.readline()),'VISUM_Tools','Tkinter.txt')
        self.f = self.path.read_text().split('\n')
        
        ##menubar
        self.menubar = tk.Menu(self.master, bg='white')
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label = "Save as...")
        self.menubar.add_cascade(label = "File", menu = self.filemenu)
        self.menubar.add_cascade(label = "Help", menu = self.filemenu)
        self.master.config(menu = self.menubar)
        
        ##master window        
        self.master.title("Changes to VISUM networks")
        self.master.geometry("520x200")
             
        ##frame1
        self.fr1 = tk.Frame(self.master)
        tk.Label(self.fr1,text = "VISUM", pady=10, relief = "groove", font='Helvetica 12 bold').grid(
            row=0, column=0, pady=3, columnspan=3, sticky=("N", "S", "E", "W")) ##geometrymanager grid
        
        #entry
        self.ent1 = tk.Entry(self.fr1, font=40, width = 40)
        self.ent1.grid(row=1,column=0)
        
        #button
        tk.Button(self.fr1, text="browse", font='Helvetica 8 bold', command = self.open_file).grid(
            row=1, column=1, padx=5, pady=5)
        tk.Button(self.fr1, text="start", font='Helvetica 8 bold', command = self.VISUM_start).grid(
            row=1, column=2, padx=5, pady=5)
        tk.Button(self.fr1, text="close \n VISUM", font='Helvetica 8 bold', command = self.VISUM_end).grid(
            row=2, column=1, sticky="wens", padx=5, pady=5)
        tk.Button(self.fr1, text="end()", font='Helvetica 8 bold', command = self.end).grid(
            row=3, column=1, sticky="wens", padx=5, pady=5)
        
        #pack Frame 1
        self.fr1.grid(row=0, column=0, padx=5, pady=5)


    def open_file(self):
        try:
            self.dat_in = fd.askopenfilename(initialdir = self.f[0],title = "Dateiauswahl",
                filetypes = [("Versionsdatei", ".ver")])
            self.ent1.insert(tk.END, self.dat_in)
        except IOError: pass
    
    def VISUM_start (self):
        self.VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
        try: self.VISUM.loadversion(self.dat_in) ##Netz,Szenrario,VISUM-Instanz
        except: pass
        
    def VISUM_end (self):
        del self.VISUM

    def end(self):
        self.master.destroy()
        
if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root)
    root.iconify()
    root.update()
    root.deiconify()
    root.mainloop()
    
del root