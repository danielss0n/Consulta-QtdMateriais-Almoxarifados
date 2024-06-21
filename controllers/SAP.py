import win32com.client
from tkinter import *
from customtkinter import *

class SAP:
    def __init__(self):
        self.session = None
        self.connect_sap()
        self.orders = []
        self.materials = []
        self.components = []
        self.data_collected_after_sap_proccess = []

    def connect_sap(self):
        try:
            sapguiauto = win32com.client.GetObject("SAPGUI")
            application = sapguiauto.GetScriptingEngine
            connection = application.Children(0)
            self.session = connection.Children(0)
        except:
            print("Abra o SAP para rodar o programa!")
            exit

    def enter_coois(self, mrp, pep):
        self.session.findById("wnd[0]").maximize
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/ncoois"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/danielsson"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-LOW").text = mrp
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PROJN-LOW").text = pep
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

    def filtrar(self):
        self.session.findById("wnd[0]/usr/cmbFILTER_BOX").setFocus
        self.session.findById("wnd[0]/usr/cmbFILTER_BOX").key = "MA_FIL"
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.position = 10
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").getAbsoluteRow(19).selected = True
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,9]").setFocus()
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,9]").caretPosition = 10
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press()
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtRANGE-LOW[2,0]").text = "1050"
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtRANGE-HIGH[3,0]").text = "1050"
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtRANGE-HIGH[3,0]").setFocus
        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtRANGE-HIGH[3,0]").caretPosition = 4
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

    def get_orders(self):
        try:
            for line in range(50):
                order = self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").getCellValue(line, "AUFNR")
                self.orders.append(order)
        except:
            pass

    def get_components(self):
        for order in self.orders:
            self.session.findById("wnd[0]").maximize
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nco03"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = order
            self.session.findById("wnd[0]/tbar[1]/btn[6]").press()
            self.get_table_values()

            self.sum_values()

    def get_table_values(self):
        self.filtrar()
        for line in range(25):
                operation = self.session.findById(f"wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-VORNR[6,{line}]").text
                operation = operation.strip()

                if operation == "____":
                    break

                if operation == "1050":
                    component = self.session.findById(f"wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2,{line}]").text
                    material = self.session.findById(f"wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1,{line}]").text
                    total_quantity = self.session.findById(f"wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,{line}]").text

                    self.data_collected_after_sap_proccess.append([component, material, total_quantity])
                    self.materials.append(material)
                

    def get_almoxarifado(self):
        print(self.materials)
        try:
            for num, material in enumerate(self.materials, start=0):
                print(self.orders)
                print(material)
                qtd_rs01 = self.find_mmbe_material(material, "RS01")
                qtd_rs02 = self.find_mmbe_material(material, "RS02")
                qtd_ia04 = self.find_mmbe_material(material, "IA04")
                self.data_collected_after_sap_proccess[num].append(qtd_rs01)
                self.data_collected_after_sap_proccess[num].append(qtd_rs02)
                self.data_collected_after_sap_proccess[num].append(qtd_ia04)
        except:
            pass


    def find_mmbe_material(self, material, almoxarifado):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nmmbe"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").text = material
        self.session.findById("wnd[0]/usr/ctxtMS_LGORT-LOW").text = almoxarifado
        
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        try:
            almoxarife = self.session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").getItemText("          4","&Hierarchy")
            almoxarife_2 = self.session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").getItemText("          5","&Hierarchy")
    
        except:
            pass
        try:
            # As vezes é na linha 4
            if "RS01" in almoxarife or "RS02" in almoxarife or "IA04" in almoxarife:
                qtd = self.session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").getItemText("          4","C          1")
                print(qtd)
                if qtd is not None or qtd != "":
                    return qtd
                else:
                    return "0"
                
            # As vezes é na linha 5
            if "RS01" in almoxarife_2 or "RS02" in almoxarife_2 or "IA04" in almoxarife_2:
                
                qtd = self.session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").getItemText("          5","C          1")
                print(qtd)
                if qtd is not None or qtd != "":
                    return qtd
                else:
                    return "0"
    
        except:
            return "0"
        
        
    def reset_components_for_next_data(self):
        self.orders = []
        self.data_collected_after_sap_proccess = []


    def sum_values(self):
        summed_components = {}
        for component in self.data_collected_after_sap_proccess:
            index = component[1]

            quantity = int(component[2].replace(',', '').strip())

            if index in summed_components:
                summed_components[index][2] = str(int(summed_components[index][2].replace(',', '').strip()) + quantity)
            else:
                summed_components[index] = component
        self.data_collected_after_sap_proccess = list(summed_components.values())