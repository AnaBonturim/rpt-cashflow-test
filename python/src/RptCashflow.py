from RptAbstract import RptAbstract

from aspose.cells import Workbook, Worksheet, CellsHelper

class RptCashflow(RptAbstract):
    
    GS_NAO_USADO_EXCECAO_LIST = [372, 331, 792]
    GS_IGNORAR_CUSTO_FORA_MATRIZ = [784, 380, 381, 794]
    TGS_IGNORAR_CUSTO_FORA_MATRIZ = [534]
    STATUS_STANDBY = [1002]
    CONNECTOR_PROJECOES = 18

    CD_TIPO_EVENTO_SCRIPT = 99
    
    def get_template_name(self):
        return "PE-CASHFLOW"
    
    def get_sheet_name(self):
        return "Cashflow Mensal"
    
    def freeze_pane(self, context, worksheet):
        
        col = CellsHelper.column_name_to_index("L")
        row = CellsHelper.row_name_to_index("9")
        
        worksheet.freeze_panes(row, col, row, col)
    
    def populate_valores(self, context):
        
        print("Populando valores.")
        
        grupo_servico_item = {}
        for item in context.get_itens():
            for cd in item.grupos_servicos:
                if not grupo_servico_item or cd not in grupo_servico_item:
                    grupo_servico_item[cd] = []
                    grupo_servico_item[cd].append(item)
        
        qt_keys = len(context.node_data_generator_list)
        group_by_list = []
        if qt_keys > 0:
            for node_data_generator in context.node_data_generator_list:
                group_by_list.append(node_data_generator.get_group_by())

        if context.crosstab_generator:
            group_by_list.append(context.crosstab_generator.get_group_by())
            
        print(f"GERANDO: {group_by_list}")
        
        group_by = ', '.join(group_by_list)
        
        if group_by:
            group_by = ', ' + group_by
            
        filter_considerar_standby = ""
        considerar_standby_parcialmente = False

        if context.cd_considerar_standby == 1:
            # Considerar parcialmente
            considerar_standby_parcialmente = True
        elif context.cd_considerar_standby == 2:
            filter_considerar_standby = f"""
                AND estudo.cd_status NOT IN ({', '.join(RptCashflow.STATUS_STANDBY)})
            """

        grupo_servico_considerar_standby_map = {}
        if considerar_standby_parcialmente:
            dez_ano = int(context.info.cd_ano_mes_previsao / 100) * 100 + 12
            self.load_grupo_servico_considerar_standby(context, grupo_servico_considerar_standby_map, dez_ano)
        
        cd_cenario = context.info.cd_cenario
        cursor = context.cursor
        
        select = f"""
            SELECT
                  a.cd_grupo_servico
                , c.cd_tipo_grupo_servico
                , a.cd_tipo_evento
                , a.cd_ano_mes
                , a.vl_evento
                , c.vl_sinal_fluxo 
                , CASE
                WHEN a.id_estudo_parceiro IS NOT NULL AND e.fl_consolidado <> 0 THEN 1
                WHEN a.id_estudo_parceiro IS NOT NULL AND e.fl_consolidado = 0 THEN 0
                WHEN d.pcParticipacao IS NULL OR a.fl_manter_100 <> 0 THEN 1
                ELSE d.pcParticipacao / 100
                END AS pcParticipacao
                , COALESCE(estudo_original.cd_status, estudo.cd_status) as cd_status
                , COALESCE(a.cd_connector, 0)
                , b.cdempreend
                {group_by}
            FROM tb_ev_estudo_evento a
            INNER JOIN tb_cenarioorcamentoempreendiqa b
                ON COALESCE(a.id_estudo_fluxo, a.id_estudo) = b.idEstudo
                    AND a.id_estudo_fase_fluxo = b.idEstudoFase
            INNER JOIN tb_ev_grupo_servico c
                ON a.cd_grupo_servico = c.cd_grupo_servico
            LEFT JOIN tb_cenarioorcamentoparticip d
                ON b.cdCenario = d.cdCenario
                    AND b.cdEmpresa = d.cdEmpresa
                    AND a.cd_ano_mes BETWEEN d.cdAnoMesDe AND d.cdAnoMesAte
            LEFT JOIN tb_ev_estudo_parceiro e
                ON a.id_estudo = e.id_estudo
                    AND a.id_estudo_parceiro = e.id_estudo_parceiro
            INNER JOIN tb_ev_estudo_fase f
                ON f.id_estudo = a.id_estudo
                    AND f.id_estudo_fase = a.id_estudo_fase_fluxo
            INNER JOIN tb_ev_estudo estudo
                ON f.id_estudo = estudo.id_estudo
            LEFT JOIN tb_ev_estudo estudo_original
                ON estudo.id_estudo_edit_original = estudo_original.id_estudo
            WHERE b.cdCenario = {cd_cenario}
            AND a.fl_rateado = 0
            {filter_considerar_standby}
        """
        
        cursor.execute(select)
        
        result = cursor.fetchall()
        
        itens = []
        keys = []
        
        qt_cols = 9
        
        for row in result:
            crosstab_key = None
            keys = []
            
            cd_grupo_servico = int(row[0])
            cd_tipo_grupo_servico = int(row[1])
            cd_tipo_evento = int(row[2])
            cd_ano_mes = int(row[3])
            vl_evento_spe = float(row[4])
            vl_sinal_fluxo = int(row[5])
            pc_participacao = float(row[6])
            cd_status = int(row[7])
            cd_connector = int(row[8])
            cd_empreend = row[9]
            
            known_cols = 9
            
            qt_cols = qt_keys + known_cols
            
            for index in range(known_cols, qt_cols):
                keys.append(row[index])
                
            if context.crosstab_generator:
                crosstab_key = row[qt_cols + 1]
            
            if pc_participacao > 0:
                pc_participacao = 1
                
            vl_evento = vl_evento_spe * pc_participacao
            
            if abs(vl_evento) < 0.01:
                continue
            
            if vl_sinal_fluxo != 0:
                vl_evento *= vl_sinal_fluxo
                
            if considerar_standby_parcialmente and cd_ano_mes >= context.info.cd_ano_mes_previsao and cd_status in RptCashflow.STATUS_STANDBY:
                grupo_servico_considerar_standby = grupo_servico_considerar_standby_map.get(cd_grupo_servico, None)
                
                if grupo_servico_considerar_standby:
                    if grupo_servico_considerar_standby.cd_empreend_incluir_list:
                        if cd_empreend not in grupo_servico_considerar_standby.cd_empreend_incluir_list:
                            continue
                    
                    if cd_ano_mes > grupo_servico_considerar_standby.am_projecoes_ate:
                        continue
                else:
                    continue
                
            if cd_ano_mes >= context.info.cd_ano_mes_previsao:
                if cd_empreend not in ['00005'] and cd_tipo_evento != RptCashflow.CD_TIPO_EVENTO_SCRIPT and cd_connector != RptCashflow.CONNECTOR_PROJECOES:
                    if cd_grupo_servico in RptCashflow.GS_IGNORAR_CUSTO_FORA_MATRIZ or cd_tipo_grupo_servico in RptCashflow.TGS_IGNORAR_CUSTO_FORA_MATRIZ:
                        continue
            
            crosstab = None

            itens.clear()

            if cd_grupo_servico > 0:
                if cd_grupo_servico in grupo_servico_item:
                    itens += grupo_servico_item[cd_grupo_servico]
            
            if itens:
                for index in range(0, qt_keys):
                    keys[index] = context.node_data_generator_list[index].convert_key(keys[index])
                
                if context.crosstab_generator:
                    crosstab_key = context.crosstab_generator.convert_key(crosstab_key)
            
            for item in itens:
                valor = vl_evento
                
                if not crosstab:
                    crosstab = context.get_crosstab(crosstab_key)                

                self.acumular(context, context.root, crosstab, valor, item)
                
                node = context.root    
                for index in range(0, len(keys)):
                    node = node.get_child(keys[index], context.node_data_generator_list[index])
                    self.acumular(context, node, crosstab, valor, item)

    def load_grupo_servico_considerar_standby(self, context, grupo_servico_considerar_standby_map, dez_ano):
        
        cursor = context.cursor
        
        select = """
            SELECT
                a.cd_grupo_servico,
                a.fl_toda_projecao,
                a.cd_empreend_incluir
            FROM tb_gs_projeto_standby a
        """
        
        cursor.execute(select)
        
        result = cursor.fetchall()
        
        for row in result:
            col = -1
            
            cd_grupo_servico = int(row[col := 1 + col])
            toda_projecao = row[col := 1 + col] != 0
            cd_empreend_incluir_list = []

            empreend_incluir = row[col := 1 + col]
            
            if empreend_incluir:
                empreend_incluir = list(empreend_incluir)
                for cd_empreend in empreend_incluir:
                    cd_empreend_incluir_list.append(cd_empreend)

            grupo_servico_considerar_standby = RptCashflow.GrupoServicoConsiderarStandby(cd_grupo_servico)
            if cd_empreend_incluir_list:
                grupo_servico_considerar_standby.cd_empreend_incluir_list = cd_empreend_incluir_list

            if not toda_projecao:
                # grupo_servico_considerar_standby.am_projecoes_ate = dez_ano
                # desconsiderando projeções do ano. alteração pedida pela Jaqueline em 27/03/2025
                grupo_servico_considerar_standby.am_projecoes_ate = 0

            grupo_servico_considerar_standby_map[cd_grupo_servico] = grupo_servico_considerar_standby

        # coloca também os aportes/saques
        grupo_servico_considerar_standby_map[710] = RptCashflow.GrupoServicoConsiderarStandby(710)
        grupo_servico_considerar_standby_map[872] = RptCashflow.GrupoServicoConsiderarStandby(872)
        
        
    class GrupoServicoConsiderarStandby:
        
        def __init__(self, cd_grupo_servico, am_projecoes_ate=999912):
            self.cd_grupo_servico = cd_grupo_servico
            self.am_projecoes_ate = am_projecoes_ate
            self.cd_empreend_incluir_list = []