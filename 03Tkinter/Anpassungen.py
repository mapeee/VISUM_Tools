# -*- coding: utf-8 -*-
"""
Created on Thu Jul 30 09:24:56 2020

@author: peter
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
        self.path = Path.joinpath(Path(self.f.readline()),'VISUM_Tools','Anpassungen.txt')
        self.f = self.path.read_text().split('\n')
        
        ##master window        
        self.master.title("Anpassungen an VISUM-Netzen")
             
        ##frame1
        self.fr1 = tk.Frame(self.master)
        tk.Label(self.fr1,text = "VISUM", pady=10, relief = "groove", font='Helvetica 12 bold').grid(
            row=0, column=0, pady=3, sticky=("N", "S", "E", "W")) ##geometrymanager grid
        tk.Label(self.fr1,text = "X", pady=10, relief = "groove", font='Helvetica 12 bold').grid(
            row=0, column=1, pady=3, sticky=("N", "S", "E", "W")) ##geometrymanager grid
        tk.Label(self.fr1,text = "Y", pady=10, relief = "groove", font='Helvetica 12 bold').grid(
            row=0, column=2, pady=3, sticky=("N", "S", "E", "W")) ##geometrymanager grid
        
        #entry
        self.ent1 = tk.Entry(self.fr1,font=40)
        self.ent1.grid(row=1,column=0)
        
        #button
        tk.Button(self.fr1, text="browse", command = self.VISUM_file).grid(
            row=1, column=1)
        tk.Button(self.fr1, text="close", command = self.end).grid(
            row=2, column=1, sticky="wens", padx=5, pady=5)
        
        #pack Frame 1
        self.fr1.grid(row=0, column=0, padx=5, pady=5)


    def VISUM_file(self):
        try:
            self.dat_in = fd.askopenfilename(initialdir = self.f[0],title = "Dateiauswahl",
                filetypes = [("Versionsdatei", ".ver")])
            self.ent1.insert(tk.END, self.dat_in)
        except IOError: pass
    
    
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