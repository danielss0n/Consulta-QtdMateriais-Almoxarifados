from customtkinter import *
import github.models.Database as db

class ShowItemWindow():
    def __init__(self, pep):
        self.pep = pep
        self.pep_data = db.get_item_by_key(self.pep)

        self.show_window()
        self.write_all_peps_to_window()
    
    def show_window(self):
        self.window = CTk()
        self.window.title(f"PEP - {self.pep}")
        self.window.geometry("600x300")

    def write_all_peps_to_window(self):
        for line, item in enumerate(self.pep_data, start=1):
            self.add_headers(item, line)
        self.window.mainloop()

    def add_headers(self, item, grid_row):
        text_field = CTkEntry(master=self.window, width=100, font=("Arial", 14), corner_radius=0)
        text_field.insert(0, "COMPONENTE")  
        text_field.grid(row=0, column=0)
        text_field = CTkEntry(master=self.window, width=100, font=("Arial", 14), corner_radius=0)
        text_field.insert(0, "MATERIAL")  
        text_field.grid(row=0, column=1)
        text_field = CTkEntry(master=self.window, width=100, font=("Arial", 14), corner_radius=0)
        text_field.insert(0, "QTD")  
        text_field.grid(row=0, column=2)
        text_field = CTkEntry(master=self.window, width=100, font=("Arial", 14), corner_radius=0)
        text_field.insert(0, "RS01")  
        text_field.grid(row=0, column=3)
        text_field = CTkEntry(master=self.window, width=100, font=("Arial", 14), corner_radius=0)
        text_field.insert(0, "RS02")  
        text_field.grid(row=0, column=4)
        text_field = CTkEntry(master=self.window, width=100, font=("Arial", 14), corner_radius=0)
        text_field.insert(0, "IA04")  
        text_field.grid(row=0, column=5)
        
        for column, value in enumerate(item, start=0):
                value = item[column]
                if value is None or value.strip() == "":
                    value = 0
                else:
                    value = value.split(",", 1)[0]
                    
                text_field = CTkEntry(master=self.window, width=100, font=("Arial", 12), corner_radius=0)
                text_field.insert(0, value)  
                text_field.grid(row=grid_row, column=column)