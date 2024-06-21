from customtkinter import *
import github.models.Database as db
import github.views.ShowItemWindow as ShowItemWindow

global_font = ("Helvetica", 25)

class ShowHistoryWindow():
    def __init__(self, window):
        self.window = CTkScrollableFrame(master=window, border_width=1, orientation="vertical", fg_color="#3b8ed0", width=300)
        self.window.pack(pady=1)
        CTkLabel(self.window, text="Histórico:", font=global_font).pack(pady=3)

        self.pep_hisotry_data = db.get_database()["consultas"]
        self.add_all_peps_to_window()
    
    def add_all_peps_to_window(self):
        for item in enumerate(self.pep_hisotry_data):
            pep_obj = item[1].keys()
            pep_number = list(pep_obj)[0]
            button = CTkButton(master=self.window, text=pep_number, command=lambda pep=pep_number: self.show_window(pep), font=global_font, corner_radius=0)
            
            button.pack()

    def show_window(self, pep_number):
        ShowItemWindow(f"{pep_number}")