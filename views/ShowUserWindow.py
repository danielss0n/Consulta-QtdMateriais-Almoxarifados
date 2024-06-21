import tkinter as tk
from tkinter import ttk
from tkinter import *
from customtkinter import *

import github.views.ShowHistoryWindow as ShowHistoryWindow
import github.controllers.StartSAP as StartSAP

global_font = ("Helvetica", 25)

# Janela inicial
class ShowUserWindow():
    def __init__(self):
        self.window = CTk()
        self.window.title("Consulta de componentes")
        self.window.geometry("500x500")
        self.window_input()
        self.window.mainloop()

    def window_input(self):
        CTkLabel(self.window, text="Consultar componentes:", font=global_font).pack()

        self.entry_pep = CTkEntry(self.window, placeholder_text="Insira o PEP...", font=global_font, width=300)
        self.entry_pep.pack(pady=10)
        self.entry_mrp = CTkEntry(self.window, placeholder_text="Insira o MRP...", font=global_font, width=300)
        self.entry_mrp.pack(pady=10)

        CTkButton(self.window, text="Consultar", command=self.get_input_values_after_click, font=global_font).pack(pady=30)
        ShowHistoryWindow(self.window)

    def get_input_values_after_click(self):
            self.pep = self.entry_pep.get()
            self.mrp = self.entry_mrp.get()

            if self.pep != "" or self.mrp != "":
                StartSAP(self.pep, self.mrp)