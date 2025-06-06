from RptCashflow import RptCashflow
from datetime import datetime

class Dados:
    
    def __init__(self, cd_cenario, agrupamento_list, cd_considerar_standby):
        self.cd_cenario = cd_cenario
        self.agrupamento_list = agrupamento_list
        self.cd_considerar_standby = cd_considerar_standby
        self.formato_tabular = False

if __name__ == "__main__":
    report = RptCashflow()
    
    dados = Dados(39822, ["regional", "empreend", "periodo"], 1)
    dados.formato_tabular = True
    
    print(f"Começando processo às {datetime.now().strftime("%H:%M:%S")}")
    report.execute(dados)
    print(f"Terminando processo às {datetime.now().strftime("%H:%M:%S")}")
