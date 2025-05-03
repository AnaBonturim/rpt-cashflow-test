from abc import ABC, abstractmethod
from datetime import date, datetime

import os
import dotenv
import psycopg2

from numbers import Number
import io
from aspose.cells import Workbook, Worksheet, Cells, CellValueType , Range, Cell, CellsException

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
        
        self.gerar_node_root(context)
        self.gerar_planilha(context, workbook)
        
        self.wrapup_workbook(context, workbook)
        self.salvar_excel(context, workbook)
        
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
                context.node_data_generator_list.append(context.get_node_data_generator(agrupamento))
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
            
            cd_tipo_grupo_servico = int(cd_tipo_grupo_servico)
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
        
        workbook = self.get_workbook(context, nm_template)
        
        if workbook is None:
            raise Exception(f"Template {nm_template} não encontrado!")
        
        nm_sheet = self.get_sheet_name()
        
        worksheet = workbook.worksheets.get(nm_sheet)
        
        if not worksheet:
            raise Exception(f"Worksheet {nm_sheet} não encontrado!")
        
        context.index_sheet_cashflow = worksheet.index
        
        cells = worksheet.cells
        
        self.populate_parametros(context, workbook, worksheet, cells)
        
        return workbook
        
    
    def get_workbook(self, context, nm_template):
        
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
    
    def populate_parametros(self, context, workbook, worksheet, cells):
        
        range_crosstab_rotulos = self.get_range(workbook, "Crosstab_Rotulos", True)
        
        if not range_crosstab_rotulos:
            raise Exception(f"Range Crosstab_Rotulos não encontrado!")
        
        col_crosstab_rotulos = range_crosstab_rotulos.first_column
        col_grupo_servico = self.get_col_range(workbook, "Filtro_GrupoServico")
        col_tipo_grupo_servico = self.get_col_range(workbook, "Filtro_TipoGrupoServico")
        
        range_row_de = range_crosstab_rotulos.first_row
        range_row_ate = range_row_de + range_crosstab_rotulos.row_count
        
        for range_row in range(range_row_de, range_row_ate):
            
            ds_linha = cells.get(range_row, col_crosstab_rotulos).string_value
            index = int(range_row - range_row_de)

            item = self.Item(index, ds_linha)
            
            if col_grupo_servico >= 0:
                cell = cells.get(range_row, col_grupo_servico)
                
                if cell and cell.value:
                    if cell.type != CellValueType.IS_NUMERIC:
                        for cd in self.parse_integer_list(cell.string_value):
                            item.add_grupo_servico(cd)
                    else:
                        item.add_grupo_servico(cell.int_value)

            if col_tipo_grupo_servico >= 0:
                cell = cells.get(range_row, col_tipo_grupo_servico)
                
                if cell and cell.value:
                    if cell.type != CellValueType.IS_NUMERIC:
                        for cd_tipo in self.parse_integer_list(cell.string_value):
                            list = context.tipo_grupo_servico_map.get(cd_tipo, None)
                            
                            if list is not None:
                                for cd in list:
                                    item.add_grupo_servico(int(cd))
                    else:
                        list = context.tipo_grupo_servico_map.get(cell.int_value, None)
                        
                        if list is not None:
                            for cd in list:
                                item.add_grupo_servico(int(cd))

            if item.has_value():
                # print(str(item))
                context.add_item(item)
    
    def gerar_node_root(self, context):
        self.populate_valores(context)
    
    def gerar_planilha(self, context, workbook):
        
        nm_sheet = self.get_sheet_name()
        
        worksheet = workbook.worksheets.get(nm_sheet)
        
        if not worksheet:
            raise Exception(f"Worksheet {nm_sheet} não encontrado!")
        
        context.index_sheet_cashflow = worksheet.index
        
        cells = worksheet.cells
        
        print("Povoando a planilha.")
        
        sorted_crosstabs = sorted(context.crosstabs.values(), key=lambda crosstab: crosstab.get_key_to_display())
        
        self.gerar_cabecalho_planilha(context, workbook)
        
        range_crosstab_rotulos = self.get_range(workbook, "Crosstab_Rotulos", True)
        crosstab_titulo = self.get_range(workbook, "Crosstab_titulo", True)
        
        col_crosstab_first = crosstab_titulo.first_column
        row_crosstab_titulo = crosstab_titulo.first_row
        qt_cols_crosstab = crosstab_titulo.column_count
        col_crosstab = col_crosstab_first - qt_cols_crosstab

        context.running_col = col_crosstab

        for crosstab in sorted_crosstabs:
            col_crosstab += qt_cols_crosstab

            if col_crosstab != col_crosstab_first:
                cells.insert_columns(col_crosstab, qt_cols_crosstab, True)
                cells.copy_columns(cells, col_crosstab_first, col_crosstab, qt_cols_crosstab)

            cell = cells.get(row_crosstab_titulo, col_crosstab)
            cell.value = crosstab.get_display()
            crosstab.apply_style(cell)

        self.hide_column(workbook, cells, "Filtro_GrupoServico")
        self.hide_column(workbook, cells, "Filtro_TipoGrupoServico")
        self.hide_column(workbook, cells, "Filtro_Partida")
        self.hide_column(workbook, cells, "Filtro_Saldo")
        self.hide_column(workbook, cells, "Partida")
        self.hide_column(workbook, cells, "Ocultar_Linha")

        range_node_descr = self.get_range(workbook, "Node_Descr", False)
        range_node_key = self.get_range(workbook, "Node_Key", False)
        range_total = self.get_range(workbook, "Crosstab_Total", True)
        
        if range_node_descr:
            context.col_node_descr = range_node_descr.first_column
            
        if range_node_key:
            context.col_node_key = range_node_key.first_column
        
        if range_crosstab_rotulos:
            self.gerar_node(context.root, context, workbook, cells, col_crosstab_first, sorted_crosstabs,
                            range_crosstab_rotulos, range_node_key, range_node_descr, range_total, 1, qt_cols_crosstab)

        # TODO
        # workbook.worksheets.remove_by_index(context.index_sheet_cashflow)

        print("Recalculando as formulas.")

        workbook.calculate_formula()
    
    def hide_column(self, workbook, cells, nm_range):
        range_col = self.get_range(workbook, nm_range)
        
        if range_col:
            last_index = range_col.column_count
            for index in reversed(range(0, last_index)):
                cells.hide_column(range_col.first_column + index)
    
    def gerar_cabecalho_planilha(self, context, workbook):
        
        template_map = {
            "cenario.cdCenario": context.info.key, 
            "cenario.dsCenario": context.info.descr,
            "agora": datetime.now().strftime("%d/%m/%y às %H:%M:%S"),
            "usuario": "mariosergioa",
            "agrupamentos": context.get_agrupamentos()
        }
        
        for range_cabecalho in workbook.worksheets.get_named_ranges(): 
            if "Cabecalho" in range_cabecalho.name:
                for index_row in range(0, range_cabecalho.row_count):
                    for index_col in range(0, range_cabecalho.column_count):
                        cell = range_cabecalho.get(index_row, index_col)
                        
                        if cell and cell.value:
                            text = cell.value
                            text = text.replace("${", "%(").replace("}", ")s") % template_map
                            
                            range_cabecalho.get(index_row, index_col).value = text
    
    def gerar_node(self, node, context, workbook, cells, col_crosstab_first, sorted_crosstabs, 
                   range_crosstab_rotulos, range_node_key, range_node_descr, range_total, level , qt_cols_crosstab):
        
        new_sheet = False
        
        if node.node_type == 'TOTAL':
            new_sheet = True
        
        if new_sheet:
            worksheet = workbook.worksheets.get(workbook.worksheets.add_copy(context.index_sheet_cashflow))
            
            try:
                worksheet.name = self.build_sheet_name(node.data.get_descr()) 
            except CellsException as e:
                print(f"ERRO: Erro ao mudar o nome da planilha para {node.data.get_descr()}: {e}")
                worksheet.name = self.build_sheet_name(node.data.get_key()) 

            cells = worksheet.cells

            context.running_row = range_crosstab_rotulos.first_row + range_crosstab_rotulos.row_count

        self.formatar_node(node, context, workbook, cells, col_crosstab_first, sorted_crosstabs,
                    range_crosstab_rotulos, range_node_key, range_node_descr, range_total,
                    level, qt_cols_crosstab)

        children_values = node.get_children_values()
        if children_values:
            generator = context.get_node_data_generator(node.children_values[0].node_type)
            sorted_children = generator.sort(children_values)

            for child in sorted_children:
                self.gerarNode(child, context, workbook, cells, col_crosstab_first, sorted_crosstabs,
                        range_crosstab_rotulos, range_node_key, range_node_descr, range_total,
                        level + 1, qt_cols_crosstab)

        if new_sheet:
            cells.delete_rows(range_crosstab_rotulos.first_row, range_crosstab_rotulos.row_count)
    
    def build_sheet_name(self, name):
        
        if not name:
            return "Sheet"
        
        for token in ['/', '\\', '[', ']', '*', '?', '-', ':']:
            name = name.replace(token, ' ')

        if len(name) > 30:
            return name[0:30]
        
        return name
        
    def formatar_node(self, node, context, workbook, cells, col_crosstab_first, sorted_crosstabs,
                    range_crosstab_rotulos, range_node_key, range_node_descr, range_total,
                    level, qt_cols_crosstab):
        
        cells.insert_rows(context.running_row, range_crosstab_rotulos.row_count)
        cells.copy_rows(cells, range_crosstab_rotulos.first_row, context.running_row, range_crosstab_rotulos.row_count)

        if range_node_key:
            index_node_key = range_node_key.first_row - range_crosstab_rotulos.first_row
            cells.get(context.running_row + index_node_key, range_node_key.first_column).value = node.data.get_key_to_display()
        
        if range_node_descr:
            index_node_descr = range_node_descr.first_row - range_crosstab_rotulos.first_row
            cell = cells.get(context.running_row + index_node_descr, range_node_descr.first_column)
            cell.value = node.data.get_descr()

            style = cell.get_style()
            style.indent_level = level - 1
            cell.set_style(style)
        
        sorted_keys = sorted(node.node_item_map)
        
        for index in sorted_keys:
            self.populate_rows(context, node, cells, sorted_crosstabs, range_total, 
                               context.running_row + index, index, col_crosstab_first, qt_cols_crosstab)

        context.running_row += range_crosstab_rotulos.row_count
    
    def populate_rows(self, context, node, cells, sorted_crosstabs, range_total, 
                               row, index, col_crosstab_first, qt_cols_crosstab):
        
        node_item = node.node_item_map.get(index, None)

        if node_item is None:
            return
        
        col_crosstab = col_crosstab_first - qt_cols_crosstab
        
        for crosstab in sorted_crosstabs:
            col_crosstab += qt_cols_crosstab
            
            valor = node_item.get_valor(crosstab)
            cell = cells.get(row, col_crosstab)
            
            if valor and valor.valor:
                cell.value = valor.valor
            else:
                cell.value = None
    
    def wrapup_workbook(self, context, workbook):
        workbook.worksheets.active_sheet_index = 0
    
    def salvar_excel(self, context, workbook):
        cd_ano_mes_previsao = context.info.cd_ano_mes_previsao
        
        workbook.save(f"../output/py-cashflow-{cd_ano_mes_previsao}.xlsx")
        
    def acumular(self, context, node, crosstab, valor, item):
        
        if valor != 0:
            nodeItem = node.get_node_item(item.index, True)
            nodeItem.add_vl_evento(valor, crosstab)
        
    def get_col_range(self, workbook, nm_range):
        
        range = self.get_range(workbook, nm_range)
        
        if not range:
            return -1
        
        return range.first_column
    
    def get_range(self, workbook, nm_range, force = False):
        
        range = workbook.worksheets.get_range_by_name(nm_range)
        
        if not range and force:
            raise Exception(f"Range {nm_range} não encontrado!")
        
        return range
    
    def parse_integer_list(self, text):
        
        list = []
        
        if text:
            valores = text.replace(' ', '').split(',')
            
            for valor in valores:
                list.append(int(valor))
            
        return list
    
    def parse_string_list(self, text):
        
        list = []
        
        if text:
            valores = text.replace(' ', '').split(',')
            
            for valor in valores:
                list.append(valor)
            
        return list
    
    class Item:
        
        def __init__(self, index, ds_linha):
            self.index = index
            self.ds_linha = ds_linha
            self.grupos_servicos = set()
        
        def __str__(self):
            return f"Item[index={self.index}, ds_linha={self.ds_linha}, grupos_servicos={self.grupos_servicos}]"
            
        def add_grupo_servico(self, cd):
            if cd:
                self.grupos_servicos.add(cd)
                
        def has_value(self):
            return bool(self.grupos_servicos)
       
    class Valor:
        
        def __init__(self):
            self.valor = 0
            
        def add_valor(self, valor):
            self.valor += valor
    
    class Node:
        
        def __init__(self, data, node_type):
            self.data = data
            self.node_type = node_type
            self.node_item_map = {}
            self.children = {}
            self.parent = None
            
        def get_children_values(self):
            return self.children.values()
        
        def get_child(self, key, generator):
            child = self.children.get(key, None)
            
            if child is None:
                child = RptAbstract.Node(generator.get_data(key), generator.get_type())
            
            return child   
            
        def get_node_item(self, index, create = False):
            node_item = self.node_item_map.get(index, None)
            
            if node_item is None and create:
                node_item = RptAbstract.NodeItem(index)
                self.node_item_map[index] = node_item
            
            return node_item
    
    class NodeItem:
        
        def __init__(self, index):
            self.index = index
            self.valores = {}
            self.crosstab_node = RptAbstract.CrosstabNode(RptAbstract.NodeDataGeneric("TOTAL", "TOTAL"), "TOTAL")
        
        
        def add_vl_evento(self, vl_evento, crosstab):
            
            valor = self.valores.get(crosstab, None)
            
            if valor is None:
                valor = RptAbstract.Valor()
                self.valores[crosstab] = valor
                
            valor.add_valor(vl_evento)
                
        def get_valor(self, crosstab):
            return self.valores.get(crosstab, None)
        
    class CrosstabNode:
        
        def __init__(self, data, node_type):
            self.data = data
            self.node_type = node_type
            self.valor = RptAbstract.Valor()
            self.children = {}
            self.parent = None
        
        def add_vl_evento(self, vl_evento):
            self.valor.add_valor(vl_evento)

        def has_child(self, key):
            return self.children and key in self.children
        
        def has_children(self):
            return self.children

        def get_child(self, key, generator, create):
            child = self.children.get(key, None)
            
            if child is None and create:
                child = RptAbstract.CrosstabNode(generator.get_data(key), generator.get_type())
                self.children[key] = child
                child.parent = self
                
            return child
        
        def get_sorted_children(self):
            return self.children.values().sort(key=lambda cn: cn.data.get_key_to_display())
    
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
    
        def sort(self, collection):
            collection.sort(key=lambda n: n.data.get_key())
    
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
    
    class NodeDataGeneric(NodeData):
        
        def __init__(self, key, descr):
            self.key = key
            self.descr = descr
            
        def get_key(self):
            return self.key
        
        def get_descr(self):
            return self.descr
        
        def get_key_to_display(self):
            return self.key
        
        def get_display(self):
            return self.descr
        
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
            style = cell.get_style()
            style.number = 17
            cell.set_style(style) 
            
    
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
                
                data = RptAbstract.PeriodoNodeData(key, date(ano, mes, 1))
                self.map[key] = data
            
            return data
        
        def get_group_by(self):
            return "a.cd_ano_mes"
        
        def convert_key(self, key):
            if isinstance(key, Number):
                return int(key)
            
            if isinstance(key, str) and key.isnumeric():
                return int(key)
            
            return key
    
    
    class Info:
        
        def __init__(self, cd_cenario, descr, cd_ano_mes_previsao):
            self.key = cd_cenario
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
            self.item_map = {}
            self.crosstabs = {}
            self.root = RptAbstract.Node(RptAbstract.NodeDataGeneric("TOTAL", "T O T A L"), "TOTAL")
            self.index_sheet_cashflow = 0
            self.running_col = 0
            self.col_node_descr = -1
            self.col_node_key = -1
            
        def add_item(self, item):
            if item:
                self.item_map[item.index] = item
                
        def get_itens(self):
            return self.item_map.values()
            
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
        
        def get_crosstab(self, crosstab_key):
            
            crosstab = self.crosstabs.get(crosstab_key, None)
            
            if crosstab is None:
                crosstab = self.crosstab_generator.get_data(crosstab_key)
                self.crosstabs[crosstab_key] = crosstab
            
            return crosstab
        
        def get_agrupamentos(self):
            agrupamentos = ""
            
            for node_data_generator in self.node_data_generator_list:
                agrupamentos += ", " + node_data_generator.get_descr()
                
            if agrupamentos:
                return agrupamentos[2:]
            else:
                return "TOTAL"
