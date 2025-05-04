
import sys

sys.path.append(r"C:\Users\abont\Projects\rpt-cashflow-test\python\src\util")

from Util import Util

class MyIterator:
    def __init__(self, data):
        self.data = data
        self.index = 0
        
    def __iter__(self):
        return self
    
    def __next__(self):
        if self.index >= len(self.data):
            raise StopIteration
        value = self.data[self.index]
        self.index += 1
        
        return value

class PeriodoIterator:
    
    def __init__(self, cd_ano_mes_min, cd_ano_mes_max):
        self.cd_ano_mes_min = cd_ano_mes_min
        self.cd_ano_mes_max = cd_ano_mes_max
        
        self.cd_ano_mes_next = self.cd_ano_mes_min
        
    def __iter__(self):
        return self
    
    def __next__(self):
        if self.cd_ano_mes_next > self.cd_ano_mes_max:
            raise StopIteration
        
        value = self.cd_ano_mes_next
        self.cd_ano_mes_next = Util.add_index_to_periodo(self.cd_ano_mes_next, 1)
        
        return value


if __name__ == "__main__":
    # iteracao = MyIterator([1, 2, 3, 4, 5])
    
    # for i in iteracao:
    #     print(f"Iteração: {i}")
        
    periodos = PeriodoIterator(202401, 202412)
    
    for periodo in periodos:
        print(f"Período {periodo}")