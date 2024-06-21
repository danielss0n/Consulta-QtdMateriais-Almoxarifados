import github.views.ShowItemWindow as ShowItemWindow
import github.controllers.SAP as sap
import github.models.Database as db

class StartSap():
    def __init__(self, pep, mrp):
        self.pep = pep
        self.mrp = mrp
        self.init_sap_proccess(self.pep, self.mrp)
        self.sap = sap
        components = self.sap.data_collected_after_sap_proccess
        item_json = {f"{pep} - {mrp}": components}

        db.save_components(item_json)
        ShowItemWindow(f"{self.pep} - {self.mrp}")
               
    def init_sap_proccess(self, pep, mrp):
        self.sap.enter_coois(mrp, pep)
        self.sap.get_orders()
        self.sap.get_components()
        self.sap.get_almoxarifado()