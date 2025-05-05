package src.main

import com.aspose.cells.*
import groovy.sql.Sql
import java.sql.Array
import java.sql.PreparedStatement
import java.sql.ResultSet

import src.main.util.*

/**
 * Created by Ana on 17/10/19.
 */
public class RptCashflow extends RptAbstract
{
    static final List<Integer> GS_NAO_USADO_EXCECAO_LIST = [372, 331, 792]
    static final List<Integer> GS_IGNORAR_CUSTO_FORA_MATRIZ = [784, 380, 381, 794] // retirei o 426 dessa lista (31/10/2024)
    static final List<Integer> TGS_IGNORAR_CUSTO_FORA_MATRIZ = [534]
    static final STATUS_STANDBY = [1002]
    static final int CONNECTOR_PROJECOES = 18

    static final int CD_TIPO_EVENTO_SCRIPT = 99

    public RptCashflow(apr)
    {
        super(apr)
    }

    @Override
    String getTemplateName(RptAbstract.Context context)
    {
        return "PE-CASHFLOW"
    }

    @Override
    String getSheetName(RptAbstract.Context context)
    {
        return "Cashflow Mensal"
    }

    @Override
    String getSheetTitle(RptAbstract.Context context)
    {
        return "Cashflow "
    }

    @Override
    String getSheetFilename(RptAbstract.Context context)
    {
        return 'Cashflow_'
    }

    @Override
    boolean isQueryGrupoServico(RptAbstract.Context context) { true }

    @Override
    boolean isQueryPrevisao(RptAbstract.Context context) { false }

    @Override
    boolean isQueryEmpreend(RptAbstract.Context context) { true }

    @Override
    def populateValores(def context)
    {
        Map<Integer, String> grupoServicoCxIndMap = populateGrupoServicoCsInd()
        List<Object> rotinaList = []
        for (Object item : context.itens.values())
        {
            if (item.rotina)
                rotinaList += item
        }
        boolean hasRotinaList = rotinaList.size() > 0

        Map<Integer, List<Object>> grupoServicoItem = [:]
        for (Object item : context.itens.values())
        {
            for (int cdGrupoServico : item.gruposServicos.keySet())
            {
                List<Object> list = grupoServicoItem[cdGrupoServico]
                if (list == null)
                {
                    list = []
                    grupoServicoItem[cdGrupoServico] = list
                }

                list << item
            }
        }

        int qtKeys = context.nodeDataGeneratorList.size()
        List<String> groupByList = []
        if (qtKeys > 0)
            for (Object nodeDataGenerator : context.nodeDataGeneratorList)
                groupByList << nodeDataGenerator.groupByIqa

        int qtCrosstabKeys = context.crosstabGeneratorList.size()
        if (qtCrosstabKeys > 0)
            for (Object nodeDataGenerator : context.crosstabGeneratorList)
                groupByList << nodeDataGenerator.groupByIqa

        if (context.crosstabGenerator)
            groupByList << context.crosstabGenerator.groupByIqa

        boolean queryEmpreend = isQueryEmpreend(context)

        Object empreendDataGenerator = null

        if (queryEmpreend)
        {
            empreendDataGenerator = context.getNodeDataGenerator('empreend')
            groupByList << empreendDataGenerator.groupByIqa
        }

        String groupBy = groupByList.join(', ')
        String filter = ""
        List<Object> filterValues = []
        for (Object filterGenerator : context.filterGeneratorList)
        {
            List<Object> filters = filterGenerator.filters
            if (filters.size() > 0)
            {
                filterValues += filters
                filter += ' AND ' + filterGenerator.groupByIqa + ' IN ('
                filter += (['?'] * filters.size()).join(', ')
                filter += ')'
            }
        }

        String periodoDe = context.periodoDe ? ' AND a.cd_ano_mes >= ? ' : ''
        String periodoAte = context.periodoAte ? ' AND a.cd_ano_mes <= ? ' : ''
        String contrapartida = !context.contrapartida ? " AND f.fl_fase_contrapartida = 0 " : ""
        String filterConsiderarStandby = ""
        boolean considerarStandbyParcialmente = false

        switch (context.cdConsiderarStandby) {
            case 1:
                //Considerar parcialmente
                considerarStandbyParcialmente = true
                break

            case 2:
                //Não considerar
                filterConsiderarStandby = """
                        AND estudo.cd_status NOT IN (${STATUS_STANDBY.join(', ')})
                    """
                break
            default:
                //Sim
                break
        }

        Map<Integer, GrupoServicoConsiderarStandby> grupoServicoConsiderarStandbyMap = [:]
        if (considerarStandbyParcialmente) {
            int dezAno = ((context.info.cdAnoMesPrevisao / 100) as int) * 100 + 12
            loadGrupoServicoConsiderarStandby context, grupoServicoConsiderarStandbyMap, dezAno
        }

        List<Integer> cenarios = context.cenarios

        String select = """
                SELECT
                    ${groupBy}
                    , a.cd_grupo_servico
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
                    , a.cd_connector
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
                WHERE b.cdCenario IN (${cenarios.join(', ')})
                AND a.fl_rateado = 0
                ${periodoDe} ${periodoAte}
                ${filter}
                ${contrapartida}
                ${filterConsiderarStandby}
            """

        // println "SELECT: $select"

        PreparedStatement ps = apr.connection.prepareStatement(select)

        int col = 1

        if (periodoDe)
            ps.setInt(col++, context.periodoDe)

        if (periodoAte)
            ps.setInt(col++, context.periodoAte)

        for (Object value : filterValues)
            ps.setObject(col++, value)

        ResultSet rs = ps.executeQuery()

        List<Object> itens = []
        Object crosstabKey
        Object[] keys = new Object[qtKeys]
        Object[] crosstabKeys = new Object[qtCrosstabKeys]

        boolean combinado = context.combinado

        Set<Integer> grupoServicoNaoUsado = new HashSet<>()

        while (rs.next())
        {
            crosstabKey = null

            col = 1
            for (int i = 0; i < qtKeys; ++i)
                keys[i] = rs.getObject(col++)

            for (int c = 0; c < qtCrosstabKeys; ++c)
                crosstabKeys[c] = rs.getObject(col++)

            if (context.crosstabGenerator)
                crosstabKey = rs.getObject(col++)

            String cdEmpreend = queryEmpreend ? rs.getString(col++) : null
            Object empreend = cdEmpreend != null ? empreendDataGenerator.getData(cdEmpreend) : null

            int cdGrupoServico = rs.getInt(col++)
            int cdTipoGrupoServico = rs.getInt(col++)
            int cdTipoEvento = rs.getInt(col++)
            int cdAnoMes = rs.getInt(col++)
            double vlEventoSPE = rs.getDouble(col++)
            double vlSinalFluxo = rs.getDouble(col++)
            double pcParticipacao = rs.getDouble(col++)
            int cdStatus = rs.getInt(col++)
            int cdConnector = rs.getInt(col++)

            if (pcParticipacao > 0)
                pcParticipacao = 1

            double vlEvento = vlEventoSPE * pcParticipacao

            if (Math.abs(vlEvento) < 0.01)
                continue

            if (vlSinalFluxo != 0)
                vlEvento *= vlSinalFluxo


            if (considerarStandbyParcialmente && cdAnoMes >= context.info.cdAnoMesPrevisao && cdStatus in STATUS_STANDBY) {
                GrupoServicoConsiderarStandby grupoServicoConsiderarStandby = grupoServicoConsiderarStandbyMap[cdGrupoServico]
                if (grupoServicoConsiderarStandby) {
                    if (grupoServicoConsiderarStandby.cdEmpreendIncluirList) {
                        if (!grupoServicoConsiderarStandby.cdEmpreendIncluirList.contains(cdEmpreend))
                            continue
                    }

                    if (cdAnoMes > grupoServicoConsiderarStandby.amProjecoesAte)
                        continue
                }
                else {
                    continue
                }
            }

            /** desativado em 05/02/2024 a pedido do Rui/Jaqueline
             //Converte aporte da linha 16.1 para a 3.12 para os custos fixos
             if (cdGrupoServico == 710 && cdEmpreend in ['00074', '00075']) {
             cdGrupoServico = 784
             }
             else if (cdGrupoServico == 547 && cdEmpreend in ['00074']) {
             cdGrupoServico = 784
             }
             */

            if (cdAnoMes >= context.info.cdAnoMesPrevisao) {
                if (!(cdEmpreend in ['00005']) && cdTipoEvento != CD_TIPO_EVENTO_SCRIPT && cdConnector != CONNECTOR_PROJECOES) {
                    if (cdGrupoServico in GS_IGNORAR_CUSTO_FORA_MATRIZ || cdTipoGrupoServico in TGS_IGNORAR_CUSTO_FORA_MATRIZ) {
                        continue
                    }
                }
            }

            def crosstab

            itens.clear()

            if (cdGrupoServico > 0) {
                if (grupoServicoItem.containsKey(cdGrupoServico)) {
                    itens += grupoServicoItem[cdGrupoServico]

                }
                else if ((!(cdGrupoServico in GS_NAO_USADO_EXCECAO_LIST)) && grupoServicoCxIndMap.containsKey(cdGrupoServico)) {
                    if (!grupoServicoNaoUsado.contains(cdGrupoServico))
                        grupoServicoNaoUsado.add(cdGrupoServico)
                }
            }

            if (hasRotinaList)
            {
                def contextRotina = [
                        cdAnoMes: cdAnoMes,
                        cdGrupoServico: cdGrupoServico,
                        cdEmpreend: cdEmpreend
                ]

                for (Object item : rotinaList)
                {
                    if (processarRotinaItem(item, contextRotina))
                    {
                        itens += item
                    }
                }
            }

            if (itens) {
                for (int i = 0; i < qtKeys; ++i)
                    keys[i] = context.nodeDataGeneratorList[i].convertKey(keys[i])

                for (int c = 0; c < qtCrosstabKeys; ++c)
                    crosstabKeys[c] = context.crosstabGeneratorList[c].convertKey(crosstabKeys[c])

                if (context.crosstabGenerator)
                    crosstabKey = context.crosstabGenerator.convertKey(crosstabKey)
            }

            for (Object item : itens)
            {
                double valor = 0

                if (!item.naPartida || cdAnoMes == context.cdAnoMesPrevisao)
                {
                    boolean use = true

                    if (item.tipoEmpreends)
                    {
                        if (empreend == null)
                            use = false
                        else if (!item.tipoEmpreends.containsKey(empreend.cdTipoEmpreend))
                            use = false
                    }

                    if (item.conditionLancto)
                    {
                        if (empreend == null)
                            use = false
                        else if (!item.conditionLancto.validate(cdAnoMes, empreend.cdAnoMesLancto))
                            use = false
                    }

                    if (item.conditionPrevisao)
                    {
                        if (!item.conditionPrevisao.validate(cdAnoMes, context.cdAnoMesPrevisao))
                            use = false
                    }

                    if (item.ignorarHolding)
                    {
                        if (cdEmpreend == '00999')
                            use = false
                    }

                    if (use)
                    {
                        valor = vlEvento
                    }
                }

                if (Math.abs(valor) < 0.01)
                    continue

                valor *= item.sinal

                if (context.crosstabAgrupavel) {
                    acumular(context, context.root, crosstabKeys, valor, item)

                    def node = context.root
                    for (int i = 0; i < keys.size(); ++i)
                    {
                        node = node.getChild(keys[i], context.nodeDataGeneratorList[i])

                        acumular(context, node, crosstabKeys, valor, item)
                    }
                }
                else {
                    if (!crosstab)
                        crosstab = context.getCrosstab(crosstabKey)

                    acumular(context, context.root, crosstab, valor, item)

                    def node = context.root
                    for (int i = 0; i < keys.size(); ++i)
                    {
                        node = node.getChild(keys[i], context.nodeDataGeneratorList[i])

                        acumular(context, node, crosstab, valor, item)
                    }
                }
            }
        }

        if (grupoServicoNaoUsado) {
            context.erro("Grupos de serviços que afetam o caixa que não foram utilizados:")
            for (int cd : grupoServicoNaoUsado)
                context.erro("\t${grupoServicoCxIndMap[cd]} (${cd})")
        }

        rs.close()
        ps.close()
    }

    @Override
    void prepareExtraContext(RptAbstract.Context context)
    {
        if (context.extraContext == null)
            context.extraContext = [:]

        context.extraContext.gerarValidador = false
    }

    @Override
    Color createCrosstabFontColorByLevel(Style style, int maxLevel, int level) {
        return style.getFont().color
    }

    @Override
    Color createCrosstabBackgroundColorByLevel(Style style, int maxLevel, int level)
    {
        int color = style.foregroundColor.toArgb()
        String hexaColor = Integer.toHexString(color)

        int alpha = Integer.parseInt(hexaColor.substring(0, 2), 16)
        int red = Integer.parseInt(hexaColor.substring(2, 4), 16)
        int green = Integer.parseInt(hexaColor.substring(4, 6), 16)
        int blue = Integer.parseInt(hexaColor.substring(6, 8), 16)

        int qtde = 30 * level
        int max = [red, green, blue].max()

        if ((max + 30 * maxLevel) <= 255) {
            red = Math.abs(red + qtde) % 255
            green = Math.abs(green + qtde) % 255
            blue = Math.abs(blue + qtde) % 255
        } else {
            red = Math.abs(red - qtde) % 255
            green = Math.abs(green - qtde) % 255
            blue = Math.abs(blue - qtde) % 255
        }

        return Color.fromArgb(alpha, red, green, blue)
    }

    @Override
    void wrapupWorkbook(Context context, Workbook workbook) {

        for (Worksheet worksheet : workbook.worksheets) {
            workbook.getWorksheets().setActiveSheetIndex(worksheet.index)
            break
        }
    }

    @Override
    void freezePane(Context context, Worksheet worksheet) {
        int[] rc = CellsHelper.cellNameToIndex("L9")
        worksheet.freezePanes(rc[0], rc[1], rc[0], rc[1])
    }

    Map<Integer, String> populateGrupoServicoCsInd() {

        Sql sql = new Sql(apr.connection)

        String select = """
            SELECT 
                a.cd_grupo_servico AS cdGrupoServico,
                a.ds_grupo_servico AS dsGrupoServico
            FROM tb_ev_grupo_servico a
            WHERE a.vl_sinal_fluxo <> 0
                AND a.fl_ignorar_ind_cashflow = 0
        """

        Map<Integer, String> result = [:]

        sql.eachRow(select) { row ->
            result[Util.iVal(row.cdGrupoServico)] = row.dsGrupoServico
        }

        return result
    }

    void loadGrupoServicoConsiderarStandby(
            def context,
            Map<Integer, GrupoServicoConsiderarStandby> grupoServicoConsiderarStandbyMap,
            int dezAno) {

        String select = """
                SELECT
                        a.cd_grupo_servico,
                        a.fl_toda_projecao,
                        a.cd_empreend_incluir
                    FROM tb_gs_projeto_standby a
            """

        PreparedStatement ps = apr.connection.prepareStatement(select)
        ResultSet rs = ps.executeQuery()

        while (rs.next()) {
            int cdGrupoServico = rs.getInt(1)
            boolean todaProjecao = rs.getInt(2) != 0
            List<String> cdEmpreendIncluirList = []

            Array arrEmpreendIncluir = rs.getArray(3)
            if (arrEmpreendIncluir) {
                def _cdEmpreendIncluirList = arrEmpreendIncluir.array
                for (String cdEmpreendIncluir : _cdEmpreendIncluirList) {
                    cdEmpreendIncluirList << cdEmpreendIncluir
                }
            }

            GrupoServicoConsiderarStandby grupoServicoConsiderarStandby = new GrupoServicoConsiderarStandby()

            grupoServicoConsiderarStandby.cdGrupoServico = cdGrupoServico
            if (cdEmpreendIncluirList.size() > 0)
                grupoServicoConsiderarStandby.cdEmpreendIncluirList = cdEmpreendIncluirList

            if (!todaProjecao) {
                grupoServicoConsiderarStandby.amProjecoesAte = dezAno
                //desconsiderando projeções do ano. alteração pedida pela Jaqueline em 27/03/2025
                grupoServicoConsiderarStandby.amProjecoesAte = 0
            }

            //alteração pedida pelo Rui em 23/01/2025 e da Jaqueline em 24/01/2025
            //if (cdGrupoServico in [251, 898, 488]) {
            //    grupoServicoConsiderarStandby.amProjecoesAte = 0
            //}

            grupoServicoConsiderarStandbyMap[cdGrupoServico] = grupoServicoConsiderarStandby
        }

        //coloca também os aportes/saques
        grupoServicoConsiderarStandbyMap[710] = new GrupoServicoConsiderarStandby(cdGrupoServico: 710)
        grupoServicoConsiderarStandbyMap[872] = new GrupoServicoConsiderarStandby(cdGrupoServico: 872)

        rs.close()
        ps.close()

    }

    static class GrupoServicoConsiderarStandby {
        int cdGrupoServico
        int amProjecoesAte = 999912
        List<String> cdEmpreendIncluirList
    }

}
