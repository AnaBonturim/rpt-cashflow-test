class Util:
    
    @staticmethod
    def add_index_to_periodo(cd_ano_mes_previsao, qt_periodos):
        
        if not qt_periodos:
            return cd_ano_mes_previsao
        
        ano = int(cd_ano_mes_previsao / 100)
        mes = cd_ano_mes_previsao % 100 + qt_periodos
        
        if mes > 12:
            while mes > 12:
                mes -= 12
                ano += 1
        elif mes < 0:
            while mes <= 0:
                mes += 12
                ano -= 1
        return ano * 100 + mes
    
# Teste
# meses = 6
# cd_ano_mes = 202408
# print(f"Adicionado {meses} meses no ano/mês {cd_ano_mes}: {Util.add_index_to_periodo(cd_ano_mes, meses)}")