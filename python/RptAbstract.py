from abc import ABC, abstractmethod
import os
import dotenv
import psycopg2
import datetime
from numbers import Number
import io
from aspose.cells import Workbook, Worksheet, Cells, CellValueType 

class RptAbstract(ABC):
    
    GENERATOR_TYPE_EMPREEND = "EMPREEND"
    GENERATOR_TYPE_PERIODO = "PERIODO"
    
    @abstractmethod
    def populate_valores(self, context):
        pass
    
    @abstractmethod
    def get_template_name(self):
        pass
    
    @abstractmethod
    def get_sheet_name(self):
        pass

    def execute(self, cd_cenario, agrupamento_list):
        
        print("Execute.")
        context = self.init(cd_cenario, agrupamento_list)
        
        workbook = self.gerar_workbook(context)
        
        # TODO: terminar o método
        
        self.populate_valores(context)
        
    def init(self, cd_cenario, agrupamento_list):
        
        cursor = self.create_cursor()
        info = self.create_info(cursor, cd_cenario)
        
        context = self.Context(cursor, info)
        
        self.populate_agrupamento(context, agrupamento_list)
        self.populate_tipo_grupo_servico(context)
        
        return context
    
    def populate_agrupamento(self, context, agrupamento_list):
        count = 0
        
        for agrupamento in agrupamento_list:
            count += 1
            if count < len(agrupamento_list):
                context.node_data_generator_list = context.get_node_data_generator(agrupamento)
            else:
                context.crosstab_generator = context.get_node_data_generator(agrupamento)
                
    def populate_tipo_grupo_servico(self, context):
        
        cursor = context.cursor
        
        select = """
            SELECT 
                cd_tipo_grupo_servico, 
                cd_grupo_servico 
            FROM tb_ev_grupo_servico
        """
        
        cursor.execute(select)
        
        result = cursor.fetchall()
        
        for row in result:
            cd_tipo_grupo_servico = row[0]
            cd_grupo_servico = row[1]
            
            if cd_tipo_grupo_servico not in context.tipo_grupo_servico_map:
                context.tipo_grupo_servico_map[cd_tipo_grupo_servico] = []
            
            context.tipo_grupo_servico_map[cd_tipo_grupo_servico].append(cd_grupo_servico)
            
        print(f"Populado {len(context.tipo_grupo_servico_map)} tipos de grupos de serviços.")
        

    def create_cursor(self):
        dotenv.load_dotenv()

        print("Buscando variáveis de ambiente.")

        HOST = os.getenv("HOST")
        DATABASE = os.getenv("DATABASE")
        USER = os.getenv("USER")
        PASSWORD = os.getenv("PASSWORD")
        
        print("Conectando com o banco de dados.")

        conn = psycopg2.connect(database=DATABASE, host=HOST, user=USER, password=PASSWORD, port="5432")
        return conn.cursor()

    def create_info(self, cursor, cd_cenario):
        
        select = """SELECT
                a.cdcenario,
                a.dscenario,
                a.cdanomesbase
            FROM tb_cenario a
            WHERE a.cdcenario = %s
        """
        
        cursor.execute(select, (cd_cenario, ))
        
        result = cursor.fetchone()
        
        cd_cenario = result[0]
        ds_cenario = result[1]
        cd_ano_mes_base = result[2]
        
        return self.Info(cd_cenario, ds_cenario, cd_ano_mes_base)
    
    
    def gerar_workbook(self, context):
        
        nm_template = self.get_template_name()
        
        template = self.get_template(context, nm_template)
        
        if template is None:
            raise Exception(f"Template {nm_template} não encontrado!")
        
        # TODO: buscar parametros
        
        
        return template
        
    
    def get_template(self, context, nm_template):
        
        cursor = context.cursor
        
        select = """
            SELECT
                by_data
            FROM tb_template
            WHERE cd_tag = 'PE-CASHFLOW'
        """
        
        cursor.execute(select)
        result = cursor.fetchone()
        
        byte_array_xlsx = result[0]
        
        return Workbook(io.BytesIO(byte_array_xlsx))   
    
    
    class NodeDataGenerator(ABC):
        
        @abstractmethod
        def get_descr(self):
            pass
        
        @abstractmethod
        def get_type(self):
            pass
        
        @abstractmethod
        def get_data(self, key):
            pass
        
        @abstractmethod
        def get_group_by(self):
            pass
        
        @abstractmethod
        def convert_key(self, key):
            pass
    
    
    class NodeData(ABC):
        
        @abstractmethod
        def get_key(self):
            pass
        
        @abstractmethod
        def get_descr(self):
            pass
        
        @abstractmethod
        def get_display(self):
            pass
        
        @abstractmethod
        def get_key_to_display(self):
            pass
        
        @abstractmethod
        def apply_style(self, cell):
            pass
    
    
    class EmpreendNodeData(NodeData):
        
        def __init__(self, cd_empreend, nm_empreend):
            self.cd_empreend = cd_empreend
            self.nm_empreend = nm_empreend
        
        def get_key(self):
            return self.cd_empreend
        
        def get_descr(self):
            return self.nm_empreend
        
        def get_display(self):
            return self.nm_empreend
        
        def get_key_to_display(self):
            return self.cd_empreend
        
        def apply_style(self, cell):
            pass
            
    
    class EmpreendGenerator(NodeDataGenerator):
        
        def __init__(self, context):
            self.map = {}
            
            self.populate(context)
            
        def populate(self, context):
            
            cd_cenario = context.info.cd_cenario
            cursor = context.cursor
            
            print(f"Buscando empreendimentos para cenário: {cd_cenario}")
            
            select = """
                SELECT
                    a.cdEmpreend,
                    COALESCE(f.nmEmpreend, a.nmEmpreend)
                FROM tb_CenarioOrcamentoEmpreend a
                LEFT JOIN tb_CenarioOrcamentoEmpreend b
                    ON a.cdCenario = b.cdCenario
                        AND a.cdEmpreendProjeto = b.cdEmpreend
                LEFT JOIN tb_Empreend f 
                    ON f.cdEmpreend = a.cdEmpreend
                WHERE a.cdCenario = %s
            """
            
            cursor.execute(select, (cd_cenario, ))
            
            result = cursor.fetchall()
            
            for row in result:
                vo = RptAbstract.EmpreendNodeData(row[0], row[1])
                self.map[vo.get_key()] = vo

            print(f"Adicionado {len(self.map)} empreendimentos no Generator")
            
        def get_descr(self):
            return "Empreendimentos"
        
        def get_type(self):
            return RptAbstract.GENERATOR_TYPE_EMPREEND
        
        def get_data(self, key):
            
            data = self.map.get(key, None)
            
            if data is None:
                
                data = RptAbstract.EmpreendNodeData(key, key)
                self.map[key] = data
            
            return data
        
        def get_group_by(self):
            return "b.cdEmpreend"
        
        def convert_key(self, key):
            return key
        
        
    class PeriodoNodeData(NodeData):
        
        def __init__(self, cd_ano_mes, dt_periodo):
            self.cd_ano_mes = cd_ano_mes
            self.dt_periodo = dt_periodo
        
        def get_key(self):
            return self.cd_ano_mes
        
        def get_descr(self):
            return self.dt_periodo
        
        def get_display(self):
            return self.dt_periodo.strftime("%b/%Y")
        
        def get_key_to_display(self):
            return self.cd_ano_mes
        
        def apply_style(self, cell):
            pass
            
    
    class PeriodoGenerator(NodeDataGenerator):
        
        def __init__(self, context):
            self.map = {}
            
            
        def get_descr(self):
            return "Períodos"
        
        def get_type(self):
            return RptAbstract.GENERATOR_TYPE_PERIODO
        
        def get_data(self, key):
            
            data = self.map.get(key, None)
            
            if data is None:
                
                ano = int(key / 100)
                mes = int(key % 100)
                
                data = RptAbstract.PeriodoNodeData(key, datetime.date(ano, mes, 1))
                self.map[key] = data
            
            return data
        
        def get_group_by(self):
            return "a.cdAnoMes"
        
        def convert_key(self, key):
            if isinstance(key, Number):
                return int(key)
            
            if isinstance(key, str) and key.isnumeric():
                return int(key)
            
            return key
    
    
    class Info:
        
        def __init__(self, cd_cenario, descr, cd_ano_mes_previsao):
            self.descr = descr
            self.cd_cenario = cd_cenario
            self.cd_ano_mes_previsao = cd_ano_mes_previsao
            
    class Context:
        
        def __init__(self, cursor, info):
            self.cursor = cursor
            self.info = info
            self.node_data_generator_list = []
            self.node_data_generator_map = {}
            self.crosstab_generator = None
            self.tipo_grupo_servico_map = {}
            
        def get_node_data_generator(self, name):
            
            node_data_generator = None
            type = name.upper()
            
            print(f"Buscando o generator para {type}")
            
            if not self.node_data_generator_map or type not in self.node_data_generator_map:
                
                if type == RptAbstract.GENERATOR_TYPE_EMPREEND:
                    node_data_generator = RptAbstract.EmpreendGenerator(self)
                elif type == RptAbstract.GENERATOR_TYPE_PERIODO:
                    node_data_generator = RptAbstract.PeriodoGenerator(self)
                else:
                    raise Exception(f"Tipo {type} não encontrado!")
                
            self.node_data_generator_map[type] = node_data_generator
                
            return node_data_generator
        
    
                
            