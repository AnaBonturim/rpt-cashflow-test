from RptAbstract import RptAbstract

class RptCashflow(RptAbstract):
    
    def get_template_name(self):
        return "PE-CASHFLOW"
    
    def get_sheet_name(self):
        return "Cashflow Mensal"
    
    def populate_valores(self, context):
        print(f"Passando pelo populate_valores: {context.info.cd_ano_mes_previsao}")
        
        
    
