package src.main

import com.aspose.cells.*

import groovy.sql.Sql
import groovy.text.SimpleTemplateEngine

import java.sql.*

import src.main.util.*

/**
 * Relatório base para cenários do PE.
 *
 * Created by mariosergioa/flavio on 16/10/15.
 */
abstract class RptAbstract
{

    Apr apr

    RptAbstract(Apr apr) {
        this.apr = apr
    }
    
    static final List<Color> BACKGROUND_COLORS = [
            Color.getBlack(),
            Color.getDimGray(),
            Color.getLightGray(),
            Color.getGhostWhite()
    ]

    static final List<Color> COLORS = [
            Color.getGhostWhite(),
            Color.getGhostWhite(),
            Color.getBlack(),
            Color.getBlack()
    ]

    static final String GENERATOR_TYPE_PROJ_CONSOLIDADO = 'PROJETO'
    static final String GENERATOR_TYPE_MACRO_PROJ_CONSOLIDADO = 'MACROPROJETO'
    static final String GENERATOR_TYPE_EMPREEND = 'EMPREEND'
    static final String GENERATOR_TYPE_REGIAO = 'REGIAO'
    static final String GENERATOR_TYPE_REGIONAL = 'REGIONAL'
    static final String GENERATOR_TYPE_SUBREGIONAL = 'SUBREGIONAL'
    static final String GENERATOR_TYPE_EMPRESA = 'EMPRESA'
    static final String GENERATOR_TYPE_EMPRESA_DIVISAO = 'EMPRESADIVISAO'
    static final String GENERATOR_TYPE_CENTRO_CUSTO = 'CENTROCUSTO'
    static final String GENERATOR_TYPE_DEPARTAMENTO = 'DEPARTAMENTO'
    static final String GENERATOR_TYPE_PERIODO = 'PERIODO'
    static final String GENERATOR_TYPE_PERIODO_ERP = 'PERIODOERP'
    static final String GENERATOR_TYPE_ANO = 'ANO'
    static final String GENERATOR_TYPE_ESTUDO_ACOMP = 'ESTUDOACOMP'
    static final String GENERATOR_TYPE_NEGOCIO = 'NEGOCIO'
    static final String GENERATOR_TYPE_GRUPO_SERVICO = 'GRUPOSERVICO'


    Context init(Dados dados)
    {
        Sql sql = new Sql(apr.connection)

        Context context = new Context()
        context.info = createInfo(sql, dados.cdCenario)
        context.cdConsiderarStandby = dados.cdConsiderarStandby
        
        context.cenarios = [dados.cdCenario]
        context.logger = new LogUtil()

        context.mensal = true
        context.formatoTabular = dados.formatoTabular
        context.abrirItens = false
        context.drilldown = true
        context.contrapartida = true
        context.abrirNiveis = true
        context.createDescrCol = true

        if (context.formatoTabular) {
            context.abrirItens = false
            context.drilldown = true
            context.abrirNiveis = true
        }

        List<String> agrupamentoList = dados.agrupamentoList
        List<String> colunas = []

        if (colunas)
        {
            context.crosstabAgrupavel = true

            for (String agrupamento : agrupamentoList)
                context.nodeDataGeneratorList << context.getNodeDataGenerator(agrupamento)

            for (String agrupamento : colunas)
                context.crosstabGeneratorList << context.getNodeDataGenerator(agrupamento)
        }
        else {
            if (agrupamentoList) {
                int count = 0
                for (String agrupamento : agrupamentoList) {
                    if (++count < agrupamentoList.size())
                        context.nodeDataGeneratorList << context.getNodeDataGenerator(agrupamento)
                    else {
                        context.crosstabGenerator = context.getNodeDataGenerator(agrupamento)
                    }
                }
            }
        }

        // context.addFilter(GENERATOR_TYPE_EMPREEND, artefato.dados?.empreends)
        // context.addFilter(GENERATOR_TYPE_PROJ_CONSOLIDADO, artefato.dados?.projetos)
        // context.addFilter(GENERATOR_TYPE_EMPRESA, artefato.dados?.empresas)
        // context.addFilter(GENERATOR_TYPE_REGIONAL, artefato.dados?.regionais)
        // context.addFilter(GENERATOR_TYPE_REGIAO, artefato.dados?.regioes)
        // context.addFilter(GENERATOR_TYPE_GRUPO_SERVICO, artefato.dados?.gruposServico)
        // context.addFilter(GENERATOR_TYPE_CENTRO_CUSTO, artefato.dados?.centrosCusto)

        prepareExtraContext(context)

        return context
    }

    Context gerarNodeRoot(Context context)
    {
        if (context == null)
            context.log("context null ")

        populateValores(context)

        if (isFillCrosstabNodeRange(context))
            context.fillCrosstabNodeRange()

        return context
    }

    Workbook gerarWorkbook(Context context, boolean bSOA)
    {
        if (context == null)
            apr.throwException("context null ")

        def byData = apr.getByData()

        if (!byData)
            apr.throwException("Template não encontrado!")

        def inputStream = new ByteArrayInputStream(byData)
        def workbook = new Workbook(inputStream)
        inputStream.close()

        String sheetName = getSheetName(context)
        Worksheet worksheet = workbook.getWorksheets().get(sheetName)
        if (worksheet == null) {
            apr.throwException("Sheet ${sheetName} não encontrado!")
        }

        context.indexSheetCashflow = worksheet.index

        Cells cells = worksheet.getCells()

        if (!bSOA)
        {
            checkFormatoTabular(context, workbook, cells)

            populateTipoGrupoServico(context)
            populateTables(context)
            populateLinhaCashflow(context)

            populateParametros(context, workbook, cells)
        }

        ajustarWorkbook context, workbook, worksheet, cells

        return workbook
    }

    void ajustarWorkbook(Context context, Workbook workbook, Worksheet worksheet, Cells cells) {

    }

    void checkFormatoTabular(Context context, Workbook workbook, Cells cells)
    {
        if (context.formatoTabular) {
            context.rcCdGroup = getRowColRange(workbook, "CDGROUP")

            if (context.createDescrCol && rcExists(context.rcCdGroup)) {
                int col = context.rcCdGroup[1] + 1
                int row = context.rcCdGroup[0]
                cells.insertColumns(col, 1)
                context.rcDsGroup = [row, col]
            }

            prepareFormatoTabular context, cells

            Range rangeNodeDescr = workbook.getWorksheets().getRangeByName("Node_Descr")
            cells.deleteRows(rangeNodeDescr.firstRow, 1, true)
        }
    }

    void execute(Dados dados)
    {
        println("Execute.")

        Context context = init(dados)
        
        Workbook workbook = gerarWorkbook(context, false)

        gerarNodeRoot(context)

        Cells cells = workbook.getWorksheets().get(context.indexSheetCashflow).getCells()
        gerarPlanilha(context, workbook, cells, context.root)
        
        wrapupWorkbook(context, workbook)

        salvarExcel (context, workbook)
    }

    void prepareExtraContext(Context context)
    {

    }

    void populateTables(Context context)
    {

    }

    boolean isRemoveFormulas()
    {
        return false
    }

    abstract def populateValores(def context)

    boolean isFillCrosstabNodeRange(Context context) { true }

    abstract String getTemplateName(Context context)
    abstract String getSheetName(Context context)
    abstract String getSheetTitle(Context context)
    abstract String getSheetFilename(Context context)

    boolean isQueryGrupoServico(Context context) { false }
    boolean isQueryEmpreend(Context context) { false }
    boolean isQueryPrevisao(Context context) { true }
    boolean isUsarSaldoFinal(Context context) { false }

    void gerarPlanilha(Context context, Workbook workbook, Cells cells, Node node) {
        context.log('Povoando planilha.')

        List<NodeData> sortedCrosstabs = null
        List<CrosstabNode> sortedCrosstabNodes = null

        if (!context.crosstabAgrupavel) {
            sortedCrosstabs = context.crosstabs.values().sort {
                crosstab ->
                    crosstab.keyToDisplay
            }
        }
        else {
            sortedCrosstabNodes = context.crosstabRoot.getSortedChildren()
        }
        
        gerarCabecalhoPlanilha context, workbook

        if (context.formatoTabular) {
            cells.ungroupRows(0, cells.maxRow, true)
            cells.unhideRows(0, cells.maxRow, 17.25)
        }

        //Oculta as linhas necessárias
        Range rangeOcultarLinha = workbook.getWorksheets().getRangeByName("Ocultar_Linha")
        Range rangeCrosstabRotulos = workbook.getWorksheets().getRangeByName("Crosstab_Rotulos")
        if (rangeOcultarLinha != null && rangeCrosstabRotulos != null) {
            for (int row = rangeCrosstabRotulos.firstRow; row < rangeCrosstabRotulos.firstRow + rangeCrosstabRotulos.rowCount; ++row) {
                if (cells.get(row, rangeOcultarLinha.firstColumn).value == 1)
                    cells.hideRow(row)
            }
        }

        Range crosstabTipo = workbook.getWorksheets().getRangeByName("Crosstab_tipo")
        int rowCrosstabTipo = crosstabTipo ? crosstabTipo.firstRow : -1

        //Duplica o crosstab pela quantidade de chaves
        Range crosstabTitulo = workbook.getWorksheets().getRangeByName("Crosstab_titulo")
        int colCrosstabFirst = crosstabTitulo.firstColumn
        int rowCrosstabTitulo = crosstabTitulo.firstRow
        int qtColsCrosstab = crosstabTitulo.columnCount
        int colCrosstab = colCrosstabFirst - qtColsCrosstab

        if (!context.mensal)
        {
            Cell cell = cells.get(rowCrosstabTitulo, colCrosstabFirst)

            def style = cell.getStyle()
            style.number = 0

            cell.setStyle(style)
        }

        if (context.cdAnoMesBaseSOA > 0)
        {
            Date dtBaseAte = toDate(context.cdAnoMesBaseSOA)

            int cdAnoDe = context.cdAnoMesBaseSOA / 100
            int cdAnoMesDe = cdAnoDe * 100 + 1

            if (cdAnoMesDe < context.info.cdAnoMesPrevisao)
                cdAnoMesDe = context.info.cdAnoMesPrevisao

            Date dtBaseDe = toDate(cdAnoMesDe)
            Date dtPrevisao = toDate(context.info.cdAnoMesPrevisao)

            Range rangeY2Dde = workbook.getWorksheets().getRangeByName("Y2D_de")
            if (rangeY2Dde)
                cells.get(rangeY2Dde.firstRow, rangeY2Dde.firstColumn).value = dtBaseDe
            Range rangeY2Date = workbook.getWorksheets().getRangeByName("Y2D_ate")
            if (rangeY2Date)
                cells.get(rangeY2Date.firstRow, rangeY2Date.firstColumn).value = dtBaseAte
            Range rangeAnoMes_PE = workbook.getWorksheets().getRangeByName("AnoMes_PE")
            if (rangeAnoMes_PE)
                cells.get(rangeAnoMes_PE.firstRow, rangeAnoMes_PE.firstColumn).value = dtPrevisao
        }

        context.runningCol = colCrosstab

        if (sortedCrosstabNodes) {
            for (CrosstabNode crosstabNode : sortedCrosstabNodes) {
                inserirColuna context, cells, crosstabNode, qtColsCrosstab, colCrosstabFirst, rowCrosstabTitulo,
                        rowCrosstabTipo, 0
            }
        }

        if (sortedCrosstabs) {
            for (NodeData crosstab : sortedCrosstabs) {
                colCrosstab += qtColsCrosstab

                if (colCrosstab != colCrosstabFirst) {
                    cells.insertColumns(colCrosstab, qtColsCrosstab, true)
                    cells.copyColumns(cells, colCrosstabFirst, colCrosstab, qtColsCrosstab)
                }

                Cell cell = cells.get(rowCrosstabTitulo, colCrosstab)
                cell.value = crosstab.display
                crosstab.applyStyle(cell)

                if (crosstabTipo) {
                    excelValue(cells, rowCrosstabTipo, colCrosstab, context.crosstabGenerator.type)
                }
            }
        }

        hideColumn(workbook, cells, "Filtro_GrupoServico")
        hideColumn(workbook, cells, "Filtro_TipoGrupoServico")
        hideColumn(workbook, cells, "Filtro_Saldo")
        hideColumn(workbook, cells, "Filtro_Partida")
        hideColumn(workbook, cells, "Filtro_Sinal")
        hideColumn(workbook, cells, "Filtro_CentroCusto")
        hideColumn(workbook, cells, "Filtro_Empreend")
        hideColumn(workbook, cells, "Filtro_Nao_Empreend")
        hideColumn(workbook, cells, "Partida")
        hideColumn(workbook, cells, "Ocultar_Linha")
        hideColumn(workbook, cells, "Filtro_Previsao")
        hideColumn(workbook, cells, "Filtro_Lancto")
        hideColumn(workbook, cells, "Filtro_TipoEmpreend")
        hideColumn(workbook, cells, "Filtro_Rotina")
        hideColumn(workbook, cells, "Filtro_IgnorarHolding")
        hideColumn(workbook, cells, "Filtro_Codigo")
        hideColumn(workbook, cells, "Filtro_Linha")

        Range rangeNodeDescr = workbook.getWorksheets().getRangeByName("Node_Descr")
        Range rangeNodeKey = workbook.getWorksheets().getRangeByName("Node_Key")
        context.rangeNodeRef = workbook.getWorksheets().getRangeByName("Node_Ref")
        Range rangeAgrupamentoTipo = workbook.getWorksheets().getRangeByName("Agrupamento_Tipo")
        Range rangeAgrupamentoCodigo = workbook.getWorksheets().getRangeByName("Agrupamento_Codigo")
        Range rangeAgrupamentoDescr = workbook.getWorksheets().getRangeByName("Agrupamento_Descr")
        Range rangeTotal = workbook.getWorksheets().getRangeByName("Crosstab_Total")

        context.colNodeDescr = rangeNodeDescr ? rangeNodeDescr.firstColumn : -1
        context.colNodeKey = rangeNodeKey ? rangeNodeKey.firstColumn : -1

        int colPartida = getColRange(workbook, "Partida")

        if (rangeCrosstabRotulos != null) {
            gerarNode(node, context, workbook, cells, colCrosstabFirst, sortedCrosstabNodes, sortedCrosstabs,
                    rangeCrosstabRotulos, rangeNodeKey, rangeNodeDescr, rangeAgrupamentoTipo,
                    rangeAgrupamentoCodigo, rangeAgrupamentoDescr, rangeTotal, 1, colPartida, qtColsCrosstab)
        }

        workbook.getWorksheets().removeAt(context.indexSheetCashflow)

        context.log('Recalculando as formulas.')

        workbook.calculateFormula()

        if (removeFormulas)
            for (int i = 2; i < workbook.getWorksheets().size(); ++i)
                workbook.getWorksheets().get(i).cells.removeFormulas()

        for (int i = 0; i < workbook.getWorksheets().count; ++i) {
            Worksheet ws = workbook.getWorksheets().get(i);
            if (i >= ws.index) {
                freezePane(context, ws)
            }
        }
    } 
    
    void gerarCabecalhoPlanilha(Context context, Workbook workbook) {
        def templateEngine = new SimpleTemplateEngine()
        def templateMap = [
                cenario     : [cdCenario : context.info.key, dsCenario : context.info.descr],
                agora       : new GregorianCalendar().format("dd/MM/yyyy 'às' HH:mm:ss"),
                usuario     : "mariosergioa",
                agrupamentos: context.agrupamentos,
                filtros     : context.filtros,
                info        : "${context.info.key} ${context.info.descr}",
                cdCenario   : context.info.key,
                dsCenario   : context.info.descr
        ] 

        for (Range rangeCabecalho : workbook.getWorksheets().getNamedRanges()) {
            if (rangeCabecalho.name.startsWith("Cabecalho")) {
                for (int indexRow = 0; indexRow < rangeCabecalho.getRowCount(); ++indexRow)
                    for (int indexCol = 0; indexCol < rangeCabecalho.getColumnCount(); ++indexCol)
                        rangeCabecalho.get(indexRow, indexCol).setValue(templateEngine.createTemplate(
                                rangeCabecalho.get(indexRow, indexCol).getStringValue()
                        ).make(templateMap))
            }
        }
    }

    void freezePane(Context context, Worksheet worksheet) {
    }

    void wrapupWorkbook(Context context, Workbook workbook) {
        workbook.getWorksheets().setActiveSheetIndex(0)
    }

    void inserirColuna(Context context, Cells cells, CrosstabNode crosstabNode, int qtColsCrosstab,
                       int colCrosstabFirst, int rowCrosstabTitulo, int rowCrosstabTipo, int level) {
        context.runningCol += qtColsCrosstab

        if (context.runningCol != colCrosstabFirst) {
            cells.insertColumns(context.runningCol, qtColsCrosstab, true)
            cells.copyColumns(cells, colCrosstabFirst, context.runningCol, qtColsCrosstab)
        }

        excelValue(cells, rowCrosstabTipo, context.runningCol, crosstabNode.nodeType)

        int maxLevel = context.crosstabGeneratorList.size() - 1

        Cell cell = cells.get(rowCrosstabTitulo, context.runningCol)
        cell.value = crosstabNode.data.display
        crosstabNode.data.applyStyle(cell)
        Style style = cell.style
        style.setForegroundColor(createCrosstabBackgroundColorByLevel(style, maxLevel, level))
        style.getFont().setColor(createCrosstabFontColorByLevel(style, maxLevel, level))
        cell.style = style

        int parentCol = context.runningCol

        for (CrosstabNode crosstabChild : crosstabNode.sortedChildren)
            inserirColuna(context, cells, crosstabChild, qtColsCrosstab, colCrosstabFirst, rowCrosstabTitulo,
                    rowCrosstabTipo, level + 1)

        if (level)
            cells.groupColumns(parentCol, context.runningCol)
    }

    void prepareFormatoTabular(Context context, Cells cells)
    {
        if (!context.formatoTabular)
            return

        int qtFormatoTabular = 0;
        if (rcExists(context.rcCdGroup)) {
            ++qtFormatoTabular
        }
        if (rcExists(context.rcDsGroup)) {
            ++qtFormatoTabular
        }

        int indexCdFormatoTabular = context.rcCdGroup[1] > context.rcDsGroup[1] && rcExists(context.rcDsGroup) ? 1 : 0
        int indexDsFormatoTabular = context.rcDsGroup[1] > context.rcCdGroup[1] && rcExists(context.rcCdGroup) ? 1 : 0
        int firstColFormatoTabular = Math.max(context.rcCdGroup[1], context.rcDsGroup[1]) - (qtFormatoTabular - 1)

        int qtKeys = context.nodeDataGeneratorList.size() - 1
        if (qtFormatoTabular > 0 && qtKeys > 0)
        {
            int qtColsFormatoTabular = qtFormatoTabular * qtKeys

            int secondColFormatoTabular = firstColFormatoTabular + qtFormatoTabular
            int rowCabecalho = Math.max(context.rcCdGroup[0], context.rcDsGroup[0])

            cells.insertColumns(secondColFormatoTabular, qtColsFormatoTabular, true)

            Range rangeFrom = cells.createRange(rowCabecalho, firstColFormatoTabular, 10, qtFormatoTabular)

            for (int index = 0; index < qtKeys; ++index) {
                int col = secondColFormatoTabular + index * qtFormatoTabular
                Range rangeTo = cells.createRange(rowCabecalho, col, 10, qtFormatoTabular)
                rangeTo.copy(rangeFrom)
            }

            //Coloca os cabeçalhos nas colunas
            for (int index = 0; index < context.nodeDataGeneratorList.size(); ++index)
            {
                NodeDataGenerator dataGenerator = context.nodeDataGeneratorList[index]
                int col = firstColFormatoTabular + index * qtFormatoTabular
                if (rcExists(context.rcCdGroup))
                    excelValue cells, context.rcCdGroup[0], col + indexCdFormatoTabular, "Cód. ${dataGenerator.descr}"

                if (rcExists(context.rcDsGroup))
                    excelValue cells, context.rcDsGroup[0], col + indexDsFormatoTabular, "Descr. ${dataGenerator.descr}"
            }
        }
        else {
            NodeDataGenerator dataGenerator = context.nodeDataGeneratorList?.getAt(0)
            if (dataGenerator && context.rcCdGroup[1] != -1)
                excelValue cells, context.rcCdGroup[0], context.rcCdGroup[1], "Cód. ${dataGenerator.descr}"

            if (dataGenerator && context.rcDsGroup[1] != -1)
                excelValue cells, context.rcDsGroup[0], context.rcDsGroup[1], "Descr. ${dataGenerator.descr}"
        }

        println("Quantidade de colunas: ${qtFormatoTabular}")
        println("index código: ${indexCdFormatoTabular}")
        println("index descr: ${indexDsFormatoTabular}")
        println("Primeira coluna: ${firstColFormatoTabular}")

        context.qtFormatoTabular = qtFormatoTabular
        context.indexCdFormatoTabular = indexCdFormatoTabular
        context.indexDsFormatoTabular = indexDsFormatoTabular
        context.firstColFormatoTabular = firstColFormatoTabular
    }

    def getFileFormatType()
    {
        return FileFormatType.XLSX
    }

    String getFileExtension()
    {
        return ".xlsx"
    }

    def salvarExcel (Context context, Workbook workbook)
    {
        context.log('Exportando.')

        workbook.save('C:\\Users\\abont\\Projects\\rpt-cashflow-test\\groovy\\output\\Cashflow.xlsx', getFileFormatType())
    }

    void buildIndice(Context context, Workbook workbook)
    {
        Worksheet worksheet = workbook.getWorksheets().get("Índice")
        if (worksheet == null)
            return

        Cells cells = worksheet.cells
        ListObject table = worksheet.getListObjects().get("Indice")
        if (table == null)
            return

        int firstRow = table.startRow
        int row = firstRow
        int colKey = table.startColumn
        int colDescr = colKey + 1
        int colSheet = colDescr + 1

        HyperlinkCollection hyperlinks = worksheet.getHyperlinks()

        for (Indice indice : context.indiceList)
        {
            String address = "'${indice.sheetName}'!A1"

            ++row
            cells.get(row, colKey).value = indice.keyToDisplay
            cells.get(row, colDescr).value = indice.descr
            cells.get(row, colSheet).value = indice.sheetName

            hyperlinks.add(row, colSheet, 1, 1, address)
        }

        table.resize(firstRow, table.startColumn, row, table.endColumn, true)
    }

    void hideColumn(Workbook workbook, Cells cells, String rangeName)
    {
        Range range = workbook.getWorksheets().getRangeByName(rangeName)
        if (range != null) {
            int lastIndex = range.columnCount
            for (int index = 0; index < lastIndex; index++)
                cells.hideColumn(range.firstColumn + index)
        }

    }

    void gerarNode(Node node, Context context, Workbook workbook, Cells cells,
                   int colCrosstabFirst, Collection<CrosstabNode> sortedCrosstabNodes,
                   Collection<NodeData> sortedCrosstabs,
                   Range rangeCrosstabRotulos, Range rangeNodeKey, Range rangeNodeDescr,
                   Range rangeAgrupamentoTipo, Range rangeAgrupamentoCodigo, Range rangeAgrupamentoDescr,
                   Range rangeTotal, int level, int colPartida, int qtColsCrosstab)
    {
        boolean newSheet = false

        if (node.nodeType == 'TOTAL' || (!context.formatoTabular && !context.drilldown && level == 2))
            newSheet = true

        if (newSheet)
        {
            Worksheet worksheet = workbook.getWorksheets().get(workbook.getWorksheets().addCopy(context.indexSheetCashflow))
            // println "SHEET NAME: ${buildSheetName(node.data?.descr?.toString())} : ${node.data?.descr} : ${node.data?.key}"
            
            try {
                worksheet.name = buildSheetName("${node.data?.descr}") 
            }
            catch (CellsException e) {
                context.erro("Erro na definição do nome da aba: ${e.getLocalizedMessage()}")
                worksheet.name = buildSheetName("${node.data?.key}")
            }
            

            cells = worksheet.getCells()

            context.runningRow = rangeCrosstabRotulos.firstRow + rangeCrosstabRotulos.rowCount

            context.indiceList << new Indice(
                    key: node.data.keyToDisplay,
                    descr: node.data.descr,
                    sheetName: worksheet.name
            )
        }

        if(!context.formatoTabular || (context.formatoTabular && node.children.size() == 0)) {
            formatarNode(node, context, workbook, cells, colCrosstabFirst, sortedCrosstabNodes, sortedCrosstabs,
                    rangeCrosstabRotulos, rangeNodeKey, rangeNodeDescr,
                    rangeAgrupamentoTipo, rangeAgrupamentoCodigo, rangeAgrupamentoDescr, rangeTotal,
                    level, colPartida, qtColsCrosstab)
        }

        if (node.childrenValues) {
            NodeDataGenerator generator = context.getNodeDataGenerator(node.childrenValues[0].nodeType)
            List<Node> sortedChildren = generator.sort(node.childrenValues)

            if (context.abrirNiveis) {
                for (Node child : sortedChildren) {
                    gerarNode(child, context, workbook, cells, colCrosstabFirst, sortedCrosstabNodes, sortedCrosstabs,
                            rangeCrosstabRotulos, rangeNodeKey, rangeNodeDescr,
                            rangeAgrupamentoTipo, rangeAgrupamentoCodigo, rangeAgrupamentoDescr, rangeTotal,
                            level + 1, colPartida, qtColsCrosstab)
                }
            }    
        }

        if (newSheet)
        {
            cells.deleteRows(rangeCrosstabRotulos.firstRow, rangeCrosstabRotulos.rowCount)
        }
    }

    String buildSheetName(String name)
    {
        if (name == null)
            return ""

        for (String token : ['/', '\\', '[', ']', '*', '?', '-', ':'])
            name = name.replace(token, ' ')
            
        // println "NAME: $name e ${name.size() > 30 ? name[0..30] : name}"

        return name.size() > 30 ? name[0..30] : name
    }

    void formatarNode(Node node, Context context, Workbook workbook, Cells cells,
                      int colCrosstabFirst, Collection<CrosstabNode> sortedCrosstabNodes, Collection<NodeData> sortedCrosstabs,
                      Range rangeCrosstabRotulos, Range rangeNodeKey, Range rangeNodeDescr,
                      Range rangeAgrupamentoTipo, Range rangeAgrupamentoCodigo, Range rangeAgrupamentoDescr,
                      Range rangeTotal, int level, int colPartida, int qtColsCrosstab)
    {
        cells.insertRows(context.runningRow, rangeCrosstabRotulos.rowCount)
        cells.copyRows(cells, rangeCrosstabRotulos.firstRow, context.runningRow, rangeCrosstabRotulos.rowCount)

        for (int index = 0; index < rangeCrosstabRotulos.rowCount; ++index)
        {
            if (rangeAgrupamentoTipo != null)
                cells.get(context.runningRow + index, rangeAgrupamentoTipo.firstColumn).value = node.nodeType
            if (rangeAgrupamentoCodigo != null)
                cells.get(context.runningRow + index, rangeAgrupamentoCodigo.firstColumn).value = node.data.keyToDisplay
            if (rangeAgrupamentoDescr != null)
                cells.get(context.runningRow + index, rangeAgrupamentoDescr.firstColumn).value = node.data.descr

            populateFormatoTabular context, node, cells, index
        }

        if (rangeNodeKey != null)
        {
            int indexNodeKey = rangeNodeKey.firstRow - rangeCrosstabRotulos.firstRow
            cells.get(context.runningRow + indexNodeKey, rangeNodeKey.firstColumn).value = node.data.keyToDisplay
        }

        if (rangeNodeDescr != null)
        {
            int indexNodeDescr = rangeNodeDescr.firstRow - rangeCrosstabRotulos.firstRow
            Cell cell = cells.get(context.runningRow + indexNodeDescr, rangeNodeDescr.firstColumn)
            cell.value = node.data.descr

            Style style = cell.style
            style.indentLevel = level - 1
            cell.style = style
        }

        int qtInsertedItensRows = 0
        def sortedKeys = node.nodeItemMap.keySet().sort({ a, b -> a <=> b })
        for (int index : sortedKeys)
        {
            List<Node> childrenUnsorted = context.abrirItens ? getChildreWithValues(node, index) : node.childrenValues.asList()

            if (!context.formatoTabular && context.abrirItens && !context.isSaldoInicial(index) && childrenUnsorted && childrenUnsorted.size() > 0)
            {
                qtInsertedItensRows += formatarNodeChildren(context, cells, colCrosstabFirst, sortedCrosstabNodes, sortedCrosstabs,
                                    rangeTotal, colPartida, qtColsCrosstab, qtInsertedItensRows, index, childrenUnsorted)
            }
            else
                populateRows(context, node, cells, sortedCrosstabNodes, sortedCrosstabs, rangeTotal,
                        context.runningRow + index, index, colPartida, colCrosstabFirst, qtColsCrosstab)
        }

        context.runningRow += rangeCrosstabRotulos.rowCount + qtInsertedItensRows
    }

    int formatarNodeChildren(Context context, Cells cells,
                             int colCrosstabFirst, Collection<CrosstabNode> sortedCrosstabNodes,
                             Collection<NodeData> sortedCrosstabs,
                             Range rangeTotal, int colPartida, int qtColsCrosstab,
                             int qtInsertedItensRows,
                             int index, List<Node> childrenUnsorted) {
        if (!childrenUnsorted || childrenUnsorted.size() == 0)
            return 0

        NodeDataGenerator generator = context.getNodeDataGenerator(childrenUnsorted[0].nodeType)
        List<Node> children = generator.sort(childrenUnsorted)
        
        int row = context.runningRow + index + qtInsertedItensRows
        cells.insertRows(row + 1, children.size())

        List<Integer> trilhaLinhas = []

        int qtRowsInsertedByChildren = 0
        int childIndex = 0
        for (Node child : children) {
            ++childIndex
            ++qtRowsInsertedByChildren

            int childRow = row + childIndex
            cells.copyRow(cells, row, childRow)
            trilhaLinhas << childRow

            alterValueWithIndent cells, childRow, context.colNodeKey, child.data.key
            alterValueWithIndent cells, childRow, context.colNodeDescr, child.data.descr

            List<Node> nextLevelChildren = getChildreWithValues(child, index)
            boolean hasChildren = nextLevelChildren && nextLevelChildren.size() > 0

            if (!hasChildren) {
                populateRows context, child, cells, sortedCrosstabNodes, sortedCrosstabs, rangeTotal,
                        childRow, index, colPartida, colCrosstabFirst, qtColsCrosstab
            }

            if (hasChildren) {
                int qtRows = formatarNodeChildren(context, cells, colCrosstabFirst, sortedCrosstabNodes, sortedCrosstabs,
                        rangeTotal, colPartida, qtColsCrosstab, qtInsertedItensRows + qtRowsInsertedByChildren,
                        index, nextLevelChildren)
                qtRowsInsertedByChildren += qtRows
                childIndex += qtRows
            }
        }

        int firstRow = row + 1
        int lastRow = row + qtRowsInsertedByChildren

        populateFormulaRow context, cells, sortedCrosstabNodes, sortedCrosstabs, row, firstRow, lastRow,
                colPartida, colCrosstabFirst, trilhaLinhas

        cells.groupRows(firstRow, lastRow)

        return qtRowsInsertedByChildren
    }

    List<Node> getChildreWithValues(Node node, int index) {
        if (!node.children || node.children.size() == 0)
            return []

        def list = []
        for (Node child : node.childrenValues) {
            NodeItem nodeItem = child.nodeItemMap[index]
            if (!nodeItem)
                continue

            if (nodeItem.valores)
                list << child
            else (nodeItem.crosstabNode.children)
                list << child
        }

        return list.unique()
    }

    void populateRows(Context context, Node node, Cells cells, Collection<CrosstabNode> sortedCrosstabNodes,
                      Collection<NodeData> sortedCrosstabs, Range rangeTotal, int row, int index, int colPartida,
                      int colCrosstabFirst, int qtColsCrosstab)
    {
        NodeItem nodeItem = node.nodeItemMap[index]

        if (!nodeItem)
            return

        if (colPartida >= 0) {
            if (nodeItem.partida.valor != 0) {
                cells.get(row, colPartida).value = nodeItem.partida.valor
            }

            formatarArrValor(node, context, cells, row, nodeItem, colPartida, nodeItem.partida)
        }

        if (sortedCrosstabs) {
            int colCrosstab = colCrosstabFirst - qtColsCrosstab
            for (NodeData crosstab : sortedCrosstabs)
            {
                colCrosstab += qtColsCrosstab
                Valor valor = nodeItem.getValor(crosstab)
                Cell cell = cells.get(row, colCrosstab)
                if (valor != null)
                {
                    if (valor.valor != 0)
                        cell.value = valor.valor
                    else
                        cell.value = null

                    formatarArrValor(node, context, cells, row, nodeItem, colCrosstab, valor)
                }
                else
                    cell.value = null
            }
        }

        if (sortedCrosstabNodes) {
            CrosstabNode crosstabNode = nodeItem.crosstabNode

            if (rangeTotal)
            {
                context.runningCol = rangeTotal.firstColumn

                Valor valor = crosstabNode.valor
                if (valor != null)
                {
                    Cell cell = cells.get(row, context.runningCol)
                    if (valor.valor != 0)
                        cell.value = valor.valor
                    else
                        cell.value = null

                    formatarArrValor(node, context, cells, row, nodeItem, context.runningCol, valor)
                }
            }

            context.runningCol = colCrosstabFirst - qtColsCrosstab

            for (CrosstabNode cabecalhoCrosstabNode : sortedCrosstabNodes)
                populateRowsRecursive context, cells, node, nodeItem, crosstabNode, cabecalhoCrosstabNode,
                        qtColsCrosstab, row
        }
    }

    void populateRowsRecursive(Context context, Cells cells, Node node, NodeItem nodeItem, CrosstabNode crosstabNode,
                               CrosstabNode cabecalhoCrosstabNode, int qtColsCrosstab, int row)
    {
        context.runningCol += qtColsCrosstab

        if (crosstabNode.hasChild(cabecalhoCrosstabNode.data.key))
        {
            crosstabNode = crosstabNode.getChild(cabecalhoCrosstabNode.data.key, null, false)

            Valor valor = crosstabNode.valor
            Cell cell = cells.get(row, context.runningCol)
            if (valor != null)
            {
                if (valor.valor != 0)
                    cell.value = valor.valor
                else
                    cell.value = null

                formatarArrValor(node, context, cells, row, nodeItem, context.runningCol, valor)
            }
            else
                cell.value = null
        }

        for (CrosstabNode cabecalhoCrosstabChild : cabecalhoCrosstabNode.sortedChildren)
            populateRowsRecursive(context, cells, node, nodeItem, crosstabNode, cabecalhoCrosstabChild, qtColsCrosstab, row)
    }

    void formatarArrValor(Node node, Context context, Cells cells, int runningRow, NodeItem nodeItem,
                          int firstCol, Valor valor) {

    }

    void populateFormulaRow(Context context, Cells cells, Collection<CrosstabNode> sortedCrosstabNodes,
                            Collection<NodeData> sortedCrosstabs, int row, int rowFirstChild,
                            int rowLastChild, int colPartida, int colCrosstabFirst, List<Integer> trilhaLinhas)
    {
        if (colPartida > -1) {
            cells.get(row, colPartida).formula = getFormulaSumRowsTrilha(trilhaLinhas, colPartida)
        }

        if (sortedCrosstabs) {
            int colCrosstab = colCrosstabFirst - 1
            for (NodeData crosstab : sortedCrosstabs)
            {
                ++colCrosstab

                Cell cell = cells.get(row, colCrosstab)
                if (cell)
                    cell.formula = getFormulaSumRowsTrilha(trilhaLinhas, colCrosstab)
                else
                    cell.value = null
            }
        }

        if (sortedCrosstabNodes) {
            context.runningCol = colCrosstabFirst - 1

            for (CrosstabNode crosstabNode : sortedCrosstabNodes)
                populateFormulaRowRecursive(context, cells, crosstabNode, row, rowFirstChild, rowLastChild, trilhaLinhas)
        }
    }

    void populateFormulaRowRecursive(Context context, Cells cells, CrosstabNode crosstabNode, int row,
                                     int rowFirstChild, int rowLastChild, List<Integer> trilhaLinhas)
    {
        context.runningCol++

        Cell cell = cells.get(row, context.runningCol)
        if (cell)
            cell.formula = getFormulaSumRowsTrilha(trilhaLinhas, context.runningCol)
        else
            cell.value = null

        for (CrosstabNode crosstabChild : crosstabNode.sortedChildren)
            populateFormulaRowRecursive(context, cells, crosstabChild, row, rowFirstChild, rowLastChild, trilhaLinhas)
    }

    String getFormulaSumRows(int firstRow, int lastRow, int col)
    {
        String colName = getColName(col)
        return "=SUM(${colName}${getRowName(firstRow)}:${colName}${getRowName(lastRow)})"
    }

    String getFormulaSumRowsTrilha(List<Integer> trilhaLinhas, int col)
    {
        int firstRow = trilhaLinhas[0]
        int lastRow = trilhaLinhas[-1]

        if (firstRow + trilhaLinhas.size() - 1 == lastRow)
            return getFormulaSumRows(firstRow, lastRow, col)

        List<String> celulas = []

        for (int row : trilhaLinhas) {
            celulas << CellsHelper.cellIndexToName(row, col)
        }

        String colName = getColName(col)
        return "=SUM(${celulas.join(',')})"
    }

    String getColName(int col)
    {
        return CellsHelper.columnIndexToName(col)
    }

    String getRowName(int row)
    {
        return CellsHelper.rowIndexToName(row)
    }

    void alterValueWithIndent(Cells cells, int row, int col, Object value)
    {
        Cell keyCell = cells.get(row, col)
        Style style = keyCell.style
        style.indentLevel = style.indentLevel + 1
        keyCell.style = style
        keyCell.value = value
    }

    void populateFormatoTabular(Context context, Node node, Cells cells, int index)
    {
        if (context.formatoTabular)
        {
            int tabIndex = context.nodeDataGeneratorList.size() - 1
            Node cnode = node

            while (cnode.parent != null)
            {
                int col = context.firstColFormatoTabular + tabIndex * context.qtFormatoTabular

                if (rcExists(context.rcCdGroup)) {
                    excelValue cells, context.runningRow + index, col + context.indexCdFormatoTabular, cnode.data.keyToDisplay
                }

                if (rcExists(context.rcDsGroup)) {
                    excelValue cells, context.runningRow + index, col + context.indexDsFormatoTabular, cnode.data.descr
                }

                tabIndex--
                cnode = cnode.parent
            }
        }
    }

    Color createCrosstabFontColorByLevel(Style style, int maxLevel, int level) {
        return COLORS[level % COLORS.size()]
    }

    Color createCrosstabBackgroundColorByLevel(Style style, int maxLevel, int level)
    {
        return BACKGROUND_COLORS[level % BACKGROUND_COLORS.size()]
    }

    static int getColRange(Workbook workbook, String name)
    {
        Range range = workbook.getWorksheets().getRangeByName(name)
        if (range != null)
            return range.getFirstColumn()
        else
            return -1
    }

    static int getRowRange(Workbook workbook, String name)
    {
        Range range = workbook.getWorksheets().getRangeByName(name)
        if (range != null)
            return range.getFirstRow()
        else
            return -1
    }

    static boolean rcExists(int[] rc) {
        return rc && (rc[0] != -1 || rc[1] != -1)
    }

    void excelValue(Cells cells, int row, int col, Object value)
    {
        if (col != -1)
            cells.get(row, col).value = value
    }

    void excelValue(Cells cells, int[] rc, Object value)
    {
        if (rc[1] != -1)
            cells.get(rc[0], rc[1]).value = value
    }

    static int[] getRowColRange(Workbook workbook, String name)
    {
        Range range = workbook.getWorksheets().getRangeByName(name)
        if (range != null)
            return [range.getFirstRow(), range.getFirstColumn()]
        else
            return [-1, -1]
    }

    static Date toDate(int cdAnoMes)
    {
        if (cdAnoMes == 0)
            return null

        return new GregorianCalendar((int)(cdAnoMes / 100), cdAnoMes % 100 - 1, 1).getTime()
    }

    void acumular(Context context, Node node, NodeData crosstab, double valor, Item item, double[] arrValor = null) {

        if (valor != 0 || arrValor != null)
        {
            NodeItem nodeItem = node.getNoteItem(item.index, true)

            if (item.naPartida)
                nodeItem.addVlEventoPartida(valor, arrValor)
            else
                nodeItem.addVlEvento(valor, crosstab, arrValor)
        }
    }

    void acumular(Context context, Node node, Object[] crosstabKeys, double valor, Item item, double[] arrValor = null)
    {
        if (valor != 0 || arrValor != null)
        {
            NodeItem nodeItem = node.getNoteItem(item.index, true)

            if (item.naPartida)
                nodeItem.addVlEventoPartida(valor, arrValor)
            else {
                CrosstabNode crosstabNode = nodeItem.crosstabNode
                crosstabNode.addVlEvento(valor, arrValor)

                for(int i = 0; i < crosstabKeys.size(); i++) {
                    crosstabNode = crosstabNode.getChild(crosstabKeys[i], context.crosstabGeneratorList[i], true)
                    crosstabNode.addVlEvento(valor, arrValor)
                }

                context.addCrosstabNode(crosstabKeys)
            }
        }
    }

    boolean processarRotinaItem(Item item, Map<String, Object> contextRotina)
    {
        return false
    }

    void populateParametros(Context context, Workbook workbook, Cells cells)
    {
        Range rangeCrosstabRotulos = workbook.getWorksheets().getRangeByName("Crosstab_Rotulos")

        if (rangeCrosstabRotulos == null)
        {
            context.log('Template sem definição de Crosstab_Rotulos')
            return
        }

        int colCrosstabRotulos = rangeCrosstabRotulos.firstColumn
        int colTag = getColRange(workbook, "Filtro_Tag")
        int colSinal = getColRange(workbook, "Filtro_Sinal")
        int colGrupoServico = getColRange(workbook, "Filtro_GrupoServico")
        int colTipoGrupoServico = getColRange(workbook, "Filtro_TipoGrupoServico")
        int colLinhaCashflow = getColRange(workbook, "Filtro_LinhaCashflow")
        int colTipoEmpreend = getColRange(workbook, "Filtro_TipoEmpreend")
        int colCentroCusto = getColRange(workbook, "Filtro_CentroCusto")
        int colEmpreendsConsiderar = getColRange(workbook, "Filtro_Empreend")
        int colEmpreendsNaoConsiderar = getColRange(workbook, "Filtro_Nao_Empreend")
        int colSaldo = getColRange(workbook, "Filtro_Saldo")
        int colPartida = getColRange(workbook, "Filtro_Partida")
        int colIgnorarHolding = getColRange(workbook, "Filtro_IgnorarHolding")
        int colLancto = getColRange(workbook, "Filtro_Lancto")
        int colPrevisao = getColRange(workbook, "Filtro_Previsao")
        int colRotina = getColRange(workbook, "Filtro_Rotina")
        int colCodigo = getColRange(workbook, "Filtro_Codigo")
        int colLinha = getColRange(workbook, "Filtro_Linha")

        int rangeRowDe = rangeCrosstabRotulos.firstRow
        int rangeRowAte = rangeRowDe + rangeCrosstabRotulos.rowCount - 1
        for (int rangeRow = rangeRowDe; rangeRow <= rangeRowAte; ++rangeRow)
        {
            String dsLinha = cells.get(rangeRow, colCrosstabRotulos).stringValue
            int index = rangeRow - rangeRowDe

            Item item = new Item(index, dsLinha)

            if (colLinhaCashflow >= 0)
            {
                Cell cell = cells.get(rangeRow, colLinhaCashflow)

                if (cell.value != null)
                {
                    for (String tagLinhaCashflow in parseStringList(cell.getStringValue()))
                    {
                        item.tagLinhaCashflow = tagLinhaCashflow
                        List<Integer> list = context.linhaCashflowMap[tagLinhaCashflow]
                        if (list != null)
                            for (int cd in list)
                                item.linhasCashflow[cd] = cd
                    }
                }
            }

            if (colGrupoServico >= 0)
            {
                Cell cell = cells.get(rangeRow, colGrupoServico)

                if (cell.value != null)
                {

                    if (cell.type != CellValueType.IS_NUMERIC)
                    {
                        for (int cd in parseList(cell.getStringValue()))
                            item.gruposServicos[cd] = cd
                    }
                    else
                    {
                        int cdGrupoServico = cell.getIntValue()
                        item.gruposServicos[cdGrupoServico] = cdGrupoServico
                    }
                }
            }

            if (colTipoGrupoServico >= 0)
            {
                Cell cell = cells.get(rangeRow, colTipoGrupoServico)

                if (cell.value != null)
                {

                    if (cell.type != CellValueType.IS_NUMERIC)
                    {
                        for (int cdTipoGrupoServico in parseList(cell.getStringValue()))
                        {
                            List<Integer> list = context.tipoGrupoServicoMap[cdTipoGrupoServico]
                            if (list != null)
                                for (int cd in list)
                                    item.gruposServicos[cd] = cd
                        }
                    }
                    else
                    {
                        int cdTipoGrupoServico = cell.getIntValue()
                        List<Integer> list = context.tipoGrupoServicoMap[cdTipoGrupoServico]
                        if (list != null)
                            for (int cd in list)
                                item.gruposServicos[cd] = cd
                    }
                }
            }

            if (colSaldo >= 0)
            {
                Cell cellSaldo = cells.get(rangeRow, colSaldo)
                if (cellSaldo.value == 'SI')
                    item.usarSaldoInicial = true
                else if (cellSaldo.value == 'SF')
                    item.usarSaldoFinal = true
            }

            if (!item.usarSaldoInicial && !item.usarSaldoFinal)
                item.usarMovimento = true

            if (colPartida >= 0)
            {
                Cell cellPartida = cells.get(rangeRow, colPartida)
                if (cellPartida.value == 1)
                    item.naPartida = true
            }

            if (colSinal >= 0)
            {
                Cell cellSinal = cells.get(rangeRow, colSinal)
                if (cellSinal.type == CellValueType.IS_NUMERIC)
                    if (cellSinal.value != 0)
                        item.sinal = cellSinal.intValue
            }

            if (colRotina >= 0)
            {
                Cell cell = cells.get(rangeRow, colRotina);
                if (cell && cell.stringValue)
                    item.rotina = cell.stringValue;
            }

            if (item.gruposServicos.size() || item.rotina || item.linhasCashflow)
                context.itens[item.index] = item
        }
    }


    static List<String> parseStringList(String text)
    {
        List<String> list = []
        
        String stripText = text.replaceAll(/[\s]+/, '')

        StringTokenizer st = new StringTokenizer(stripText, ",")
        while (st.hasMoreTokens())
        {
            String cd = st.nextToken()
            list << cd as String
        }
        
        return list
    }

    static List<Integer> parseIntegerList(String text)
    {
        List<Integer> list = []

        String stripText = text.replaceAll(/[\s]+/, '')

        StringTokenizer st = new StringTokenizer(stripText, ",")
        while (st.hasMoreTokens())
        {
            int cd = Integer.parseInt(st.nextToken())
            list << cd
        }

        return list
    }

    static List<Integer> parseList(String text)
    {
        List<Integer> list = []

        StringTokenizer st = new StringTokenizer(text, ", ")
        while (st.hasMoreTokens())
        {
            int cd = Integer.parseInt(st.nextToken())
            list << cd
        }

        return list
    }

    Info createInfo(Sql sql, int key)
    {
        
        return getInfo(sql, key) {
            """
                SELECT
                        c.cdAnoMesPrevisao AS "cdAnoMesPrevisao",
                        c.dscenario AS "descr"
                FROM tb_cenario c
                WHERE c.cdcenario = :key
            """
        }
    }

    Info getInfo(Sql sql, int key, Closure<String> createSelect)
    {
        Info info = new Info( key : key)

        String select = createSelect()

        def row = sql.firstRow(select, [key : key])

        info.cdAnoMesPrevisao = row.cdAnoMesPrevisao
        info.descr = row.descr

        if (!info.cdAnoMesPrevisao)
            apr.throwException("ID ${key} não encontrado!")

        return info
    }


    void populateTipoGrupoServico(Context context)
    {
        PreparedStatement ps = apr.connection.prepareStatement("SELECT cd_tipo_grupo_servico, cd_grupo_servico FROM tb_ev_grupo_servico")

        ResultSet rs = ps.executeQuery()

        while (rs.next())
        {
            int cdTipoGrupoServico = rs.getInt(1)
            int cdGrupoServico = rs.getInt(2)

            List<Integer> list = context.tipoGrupoServicoMap[cdTipoGrupoServico]
            if (list == null)
            {
                list = []
                context.tipoGrupoServicoMap[cdTipoGrupoServico] = list
            }

            list << cdGrupoServico
        }

        rs.close()
        ps.close()
    }

    void populateLinhaCashflow(Context context)
    {
        String query = "SELECT cd_linha_cashflow, tag_linha_cashflow FROM tb_ev_linha_cashflow"

        try (PreparedStatement ps = apr.connection.prepareStatement(query))
        {
            try (ResultSet rs = ps.executeQuery())
            {
                while (rs.next())
                {
                    int cdLinhaCashflow = rs.getInt(1)
                    String tagLinhaCashflow = rs.getString(2)

                    List<Integer> list = context.linhaCashflowMap[tagLinhaCashflow]
                    if (list == null)
                    {
                        list = []
                        context.linhaCashflowMap[tagLinhaCashflow] = list
                    }

                    list << cdLinhaCashflow
                }
            }
        }
    }

    static class Valor
    {
        double valor
        double[] arrValor

        public void addValores(double valor, double[] arrValor) {
            this.valor += valor

            if (arrValor != null) {
                if (this.arrValor == null) {
                    this.arrValor = new double[arrValor.length]
                }

                for (int i = 0; i < arrValor.length; ++i) {
                    this.arrValor[i] += arrValor[i]
                }
            }
        }
    }

    static class Item
    {
        int index
        String dsLinha
        String tagLinhaCashflow
        String rotina
        Map<Integer, Integer> gruposServicos = [:]
        Map<Integer, Integer> linhasCashflow = [:]
        Map<String, String> contas = [:]
        Map<String, String> contasCr = [:]
        Map<String, String> tags = [:]
        Map<Integer, Integer> tipoEmpreends = [:]
        Map<String, String> centrosCusto = [:]
        Map<String, String> empreendsConsiderar
        Map<String, String> empreendsNaoConsiderar
        Map<String, String> codigos = [:]
        Map<Integer, Integer> linhas = [:]
        boolean usarSaldoInicial
        boolean usarMovimento
        boolean usarSaldoFinal
        boolean naPartida
        boolean ignorarHolding
        int sinal = 1
        Condition conditionLancto
        Condition conditionPrevisao

        Item(int index, String dsLinha)
        {
            this.index = index
            this.dsLinha = dsLinha
        }

    }

    static class NodeItem
    {
        int index
        Map<NodeData, Valor> valores = [:]
        Valor partida = new Valor()

        CrosstabNode crosstabNode

        public NodeItem(int index)
        {
            this.index = index
            this.crosstabNode = new CrosstabNode(new NodeDataGeneric(key: "TOTAL", descr: "TOTAL"), "TOTAL")
        }

        void addVlEventoPartida(double vlEvento) {
            addVlEventoPartida(vlEvento, null)
        }

        void addVlEventoPartida(double vlEvento, double[] arrValor)
        {
            partida.addValores(vlEvento, arrValor)
        }

        void addVlEvento(double vlEvento, NodeData crosstab, double[] arrValor = null)
        {
            Valor valor = valores[crosstab]
            if (valor == null)
            {
                valor = new Valor()
                valores[crosstab] = valor
            }

            valor.addValores(vlEvento, arrValor)
        }

        Valor getValor(NodeData crosstab)
        {
            return valores[crosstab]
        }
    }

    static class Node
    {
        NodeData data
        String nodeType

        public Node(NodeData data, String nodeType)
        {
            this.data = data
            this.nodeType = nodeType
        }

        /**
         * key = item index
         */
        Map<Integer, NodeItem> nodeItemMap = [:]

        Map<Object, Node> children = [:]
        Node parent

        Node()
        {
        }

        Collection<Node> getChildrenValues() {
            return children?.values()
        }
        
        Node getChild(Object key, NodeDataGenerator generator)
        {
            Node child = children[key]
            if (child == null)
            {
                child = new Node(generator.getData(key), generator.type)
                children[key] = child
                child.parent = this
            }

            return child
        }

        NodeItem getNoteItem(int index, boolean create = false)
        {
            NodeItem nodeItem = nodeItemMap[index]
            if (nodeItem == null && create)
            {
                nodeItem = new NodeItem(index)
                nodeItemMap[index] = nodeItem
            }

            return nodeItem
        }

        Collection<NodeData> getCrosstabsForNode()
        {
            Map<Object, NodeData> map = [:]

            for (NodeItem nodeItem : nodeItemMap.values())
                for (NodeData crosstab : nodeItem.valores.keySet())
                    if (!map.containsKey(crosstab.key))
                        map[crosstab.key] = crosstab

            return map.values()
        }
    }

    static class CrosstabNode
    {
        NodeData data
        String nodeType
        Valor valor

        Map<Object, CrosstabNode> children = [:]
        CrosstabNode parent

        public CrosstabNode(NodeData data, String nodeType)
        {
            this.data = data
            this.nodeType = nodeType
            this.valor = new Valor()
        }

        void addVlEvento(double vlEvento, double[] arrValor) {
            valor.addValores(vlEvento, arrValor)
        }

        boolean hasChild(Object key) {
            return this.children.containsKey(key)
        }

        boolean hasChildren() {
            return this.children
        }

        CrosstabNode getChild(Object key, NodeDataGenerator generator, boolean create)
        {
            CrosstabNode child = children[key]
            if (child == null && create)
            {
                child = new CrosstabNode(generator.getData(key), generator.type)
                children[key] = child
                child.parent = this
            }

            return child
        }

        Collection<CrosstabNode> getSortedChildren() {
            return this.children.values().sort({ cn -> cn.data.keyToDisplay })
        }
    }

    static class Empresa implements NodeData
    {
        int cdEmpresa
        String dsEmpresa
        String cdEmpresaERP

        @Override
        Object getKey()
        {
            return cdEmpresa
        }

        @Override
        String getDescr()
        {
            return dsEmpresa?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return dsEmpresa?.toUpperCase()
        }

        @Override
        String getKeyToDisplay()
        {
            return "${cdEmpresa}"
        }

        @Override
        void applyStyle(Cell cell)
        {

        }

        boolean equals(o)
        {
            if (this.is(o)) return true
            if (getClass() != o.class) return false

            Empresa that = (Empresa) o

            if (cdEmpresa != that.cdEmpresa) return false

            return true
        }

        int hashCode()
        {
            return cdEmpresa
        }

    }

    static class Empreend implements NodeData
    {
        String cdEmpreend
        String nmEmpreend
        int cdAnoMesLancto
        int cdAnoMesInicioConstrucao
        int cdAnoMesChaves
        String cdEmpreendProjeto
        String nmEmpreendProjeto
        int cdRegional
        String dsRegional
        int cdSubRegional
        String dsSubRegional
        int cdRegiao
        String dsRegiao
        int cdTipoEmpreend

        @Override
        Object getKey()
        {
            return cdEmpreend
        }

        @Override
        String getDescr()
        {
            return nmEmpreend?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdEmpreend} ${nmEmpreend?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {

        }
    }

    static class NodeDataGeneric implements NodeData
    {
        Object key
        String descr

        @Override
        Object getDisplay()
        {
            return descr?.toUpperCase()
        }

        @Override
        String getKeyToDisplay()
        {
            return key
        }

        @Override
        void applyStyle(Cell cell)
        {

        }
    }

    static class NegocioNodeData implements NodeData
    {
        int cdNegocio
        String dsNegocio

        @Override
        Object getKey()
        {
            return cdNegocio
        }

        @Override
        String getDescr()
        {
            return dsNegocio?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdNegocio}${dsNegocio?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return cdNegocio
        }

        @Override
        void applyStyle(Cell cell)
        {

        }

        boolean equals(o)
        {
            if (this.is(o)) return true
            if (getClass() != o.class) return false

            NegocioNodeData that = (NegocioNodeData) o

            if (key != that.key) return false

            return true
        }

        int hashCode()
        {
            return (key != null ? key.hashCode() : 0)
        }
    }

    class NegocioGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, NegocioNodeData> map = [:]

        public NegocioGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new NegocioNodeData(
                        cdNegocio: convertKey(key),
                        dsNegocio: "Sem categorização"
                )
                map[key] = data
            }

            return data
        }

        void populate(Context context)
        {
            PreparedStatement ps = apr.connection.prepareStatement("""
                    SELECT
                          a.cd_negocio
                        , a.ds_negocio
                        FROM tb_ev_negocio a
                """)

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                NegocioNodeData vo = new NegocioNodeData(
                        cdNegocio: rs.getInt(1),
                        dsNegocio: rs.getString(2)
                )

                map[vo.cdNegocio] = vo
            }

            rs.close()
            ps.close()
        }

        @Override
        String getDescr()
        {
            return "Negócio"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_NEGOCIO
        }

        @Override
        String getGroupBy()
        {
            return "e.cd_negocio"
        }

        @Override
        String getGroupByIqa()
        {
            return "e.cd_negocio"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdNegocio"
        }

        @Override
        Object convertKey(Object key)
        {
            return Util.iVal(key)
        }

    }

    class EstudoAcompGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, NodeDataGeneric> map = [:]

        public EstudoAcompGenerator(Context context)
        {
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new NodeDataGeneric(
                        key: convertKey(key),
                        descr: key?.toString()
                )

                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Estudo/Acompanhamento"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_ESTUDO_ACOMP
        }

        @Override
        String getGroupBy()
        {
            return """
                CASE
                    WHEN b.cdanomeslancto >= c.cdAnoMesPrevisao THEN 'A Lançar'
                    WHEN b.cdAnoMesInicioConstrucao >= c.cdanomesprevisao THEN 'Lançado'
                    WHEN b.cdanomeschaves >= c.cdanomesprevisao THEN 'Andamento'
                    ELSE 'Entregue'
                END
            """
        }

        @Override
        String getGroupByIqa()
        {
            return """
                CASE
                    WHEN b.cdanomeslancto >= c.cdAnoMesPrevisao THEN 'A Lançar'
                    WHEN b.cdAnoMesInicioConstrucao >= c.cdanomesprevisao THEN 'Lançado'
                    WHEN b.cdanomeschaves >= c.cdanomesprevisao THEN 'Andamento'
                    ELSE 'Entregue'
                END
            """
        }
        
        @Override
        String getGroupByVO()
        {
            //Nesse caso é necessário criar no mapa o racional do getGroupBy com um atributo que tenha o nome de tipoEstudo
            return "tipoEstudo"
        }

        @Override
        Object convertKey(Object key)
        {
            return key.toString()
        }
    }

    static class PeriodoNodeData implements NodeData, DateFilter
    {
        static final String FORMAT = "MMM/yyyy"

        int cdAnoMes
        DateTime dtPeriodo

        @Override
        Object getKey()
        {
            return cdAnoMes
        }

        @Override
        String getDescr()
        {
            return dtPeriodo.toDate().format(FORMAT)
        }

        @Override
        Object getDisplay()
        {
            return dtPeriodo.toDate()
        }

        @Override
        String getKeyToDisplay()
        {
            return dtPeriodo
        }

        @Override
        void applyStyle(Cell cell)
        {
            Style style = cell.style
            style.number = 17
            cell.style = style
        }

        boolean equals(o)
        {
            if (this.is(o)) return true
            if (getClass() != o.class) return false

            PeriodoNodeData that = (PeriodoNodeData) o

            if (key != that.key) return false

            return true
        }

        int hashCode()
        {
            return (key != null ? key.hashCode() : 0)
        }

        @Override
        boolean isValidByAnoMes(NodeDataGenerator generator, Object key) {
            if (generator instanceof PeriodoGenerator)
                return this.cdAnoMes == key
            else if (generator instanceof AnoGenerator)
                return Util.iVal(this.cdAnoMes / 100) == key

            return false
        }
    }

    static class GrupoServico implements NodeData
    {
        int cdGrupoServico
        String dsGrupoServico

        @Override
        Object getKey()
        {
            return cdGrupoServico
        }

        @Override
        String getDescr()
        {
            return dsGrupoServico?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdGrupoServico} ${dsGrupoServico?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }

    }

    class GrupoServicoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, GrupoServico> map = [:]

        public GrupoServicoGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new GrupoServico(
                        cdGrupoServico: key,
                        dsGrupoServico: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Grupos de Serviços"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_GRUPO_SERVICO
        }

        @Override
        String getGroupBy()
        {
            return "c.cd_grupo_servico"
        }

        @Override
        String getGroupByIqa()
        {
            return "c.cd_grupo_servico"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdGrupoServico"
        }

        @Override
        Object convertKey(Object key)
        {
            if (key instanceof Integer)
                return key

            return Integer.parseInt(key?.toString())
        }

        void populate(Context context)
        {
            PreparedStatement ps = apr.connection.prepareStatement("SELECT cd_grupo_servico, ds_grupo_servico FROM tb_ev_grupo_servico")

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                GrupoServico vo = new GrupoServico(
                        cdGrupoServico: rs.getInt(1),
                        dsGrupoServico: rs.getString(2)
                )

                map[vo.cdGrupoServico] = vo
            }

            rs.close()
            ps.close()
        }

    }

    interface FillNodeRange {
        Iterator<Object> getKeys()
    }

    interface DateFilter {
        boolean isValidByAnoMes(NodeDataGenerator generator, Object key)
    }

    class PeriodoIterator implements Iterator<Object>
    {
        int mincdAnoMes = 0
        int maxcdAnoMes = 0
        int nextcdAnoMes = 0

        PeriodoIterator(int mincdAnoMes, int maxcdAnoMes) {
            // println "Mínimo e máximo: ${mincdAnoMes} e ${maxcdAnoMes}"
            this.mincdAnoMes = mincdAnoMes
            this.maxcdAnoMes = maxcdAnoMes

            nextcdAnoMes = mincdAnoMes
        }

        @Override
        boolean hasNext() {
            return nextcdAnoMes <= maxcdAnoMes
        }

        @Override
        Object next() {
            int value = nextcdAnoMes
            nextcdAnoMes = Util.addIndexToPeriodo(nextcdAnoMes, 1)

            return value
        }

    }

    class PeriodoGenerator extends AbstractDataGenerator implements NodeDataGenerator, FillNodeRange
    {
        Map<Object, PeriodoNodeData> map = [:]
        Integer mincdAnoMes = null
        Integer maxcdAnoMes = null

        PeriodoGenerator(Context context)
        {
        }

        @Override
        Iterator<Object> getKeys() {
            return new PeriodoIterator(mincdAnoMes, maxcdAnoMes)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new PeriodoNodeData(
                        cdAnoMes: convertKey(key),
                        dtPeriodo: AsposeCellsHelper.amToDateTime(key)
                )

                map[data.cdAnoMes] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Períodos"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_PERIODO
        }

        @Override
        String getGroupBy()
        {
            return "a.cdAnoMes"
        }

        @Override
        String getGroupByIqa()
        {
            return "a.cd_ano_mes"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdAnoMes"
        }

        @Override
        Object convertKey(Object key)
        {
            int cdAnoMes = Util.iVal(key)

            if (cdAnoMes != 0) {
                if (mincdAnoMes == null || cdAnoMes < mincdAnoMes)
                    mincdAnoMes = cdAnoMes

                if (maxcdAnoMes == null || cdAnoMes > maxcdAnoMes)
                    maxcdAnoMes = cdAnoMes
            }

            return cdAnoMes
        }

    }

    static class AnoNodeData implements NodeData, DateFilter
    {
        int ano

        @Override
        Object getKey()
        {
            return ano
        }

        @Override
        String getDescr()
        {
            return ano
        }

        @Override
        Object getDisplay()
        {
            return ano
        }

        @Override
        String getKeyToDisplay()
        {
            return ano
        }

        @Override
        void applyStyle(Cell cell)
        {
            Style style = cell.style
            style.number = 1
            cell.style = style
        }

        boolean equals(o)
        {
            if (this.is(o)) return true
            if (getClass() != o.class) return false

            AnoNodeData that = (AnoNodeData) o

            if (key != that.key) return false

            return true
        }

        int hashCode()
        {
            return (key != null ? key.hashCode() : 0)
        }

        @Override
        boolean isValidByAnoMes(NodeDataGenerator generator, Object key) {
            if (generator instanceof PeriodoGenerator)
                return this.ano == Util.iVal(key / 100)
            else if (generator instanceof AnoGenerator)
                return this.ano == key

            return false
        }
    }

    class AnoIterator implements Iterator<Object>
    {
        int minCdAno = 0
        int maxCdAno = 0
        int nextCdAno = 0

        AnoIterator(int minCdAno, int maxCdAno) {
            this.minCdAno = minCdAno
            this.maxCdAno = maxCdAno

            nextCdAno = minCdAno
        }

        @Override
        boolean hasNext() {
            return nextCdAno <= maxCdAno
        }

        @Override
        Object next() {
            int value = nextCdAno

            ++nextCdAno

            return value
        }

    }

    class AnoGenerator extends AbstractDataGenerator implements NodeDataGenerator, FillNodeRange
    {
        Map<Object, AnoNodeData> map = [:]
        Integer minCdAno = null
        Integer maxCdAno = null

        public AnoGenerator(Context context)
        {
        }

        @Override
        Iterator<Object> getKeys() {
            return new AnoIterator(minCdAno, maxCdAno)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new AnoNodeData(
                        ano: convertKey(key)
                )

                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Anos"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_ANO
        }

        @Override
        String getGroupBy()
        {
            return "a.cdAno"
        }

        @Override
        String getGroupByIqa()
        {
            return "(a.cd_ano_mes / 100)::INTEGER"
        }
        
        @Override
        String getGroupByVO()
        {
            //Nesse caso é necessário criar no mapa o racional do getGroupBy com um atributo que tenha o nome de ano
            return "ano"
        }

        @Override
        Object convertKey(Object key)
        {
            int cdAno = Util.iVal(key)

            if (cdAno != 0) {
                if (minCdAno == null || cdAno < minCdAno)
                    minCdAno = cdAno

                if (maxCdAno == null || cdAno > maxCdAno)
                    maxCdAno = cdAno
            }

            return cdAno
        }
    }

    class EmpresaGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Empresa> map = [:]

        public EmpresaGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Empresa(
                        cdEmpresa: key,
                        cdEmpresaERP: key?.toString(),
                        dsEmpresa: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Empresas"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_EMPRESA
        }

        @Override
        String getGroupBy()
        {
            return "a.cdEmpresa"
        }

        @Override
        String getGroupByIqa()
        {
            return "a.cd_empresa"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdEmpresa"
        }

        @Override
        Object convertKey(Object key)
        {
            return Util.iVal(key)
        }

        void populate(Context context)
        {
            List<Integer> cenarios = context.cenarios

            PreparedStatement ps = apr.connection.prepareStatement("""
                    SELECT
                          a.cdEmpresa
                        , a.dsEmpresa
                        , b.cdEmpresaErp
                        FROM tb_cenarioorcamentoempresa a
                            LEFT JOIN tb_empresa b
                                ON a.cdEmpresa = b.cdEmpresa
                        WHERE a.cdCenario IN (${cenarios.join(', ')})
                """)

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Empresa vo = new Empresa(
                        cdEmpresa: rs.getInt(1),
                        dsEmpresa: rs.getString(2),
                        cdEmpresaERP: rs.getString(3)
                )

                map[vo.cdEmpresa] = vo
            }

            rs.close()
            ps.close()
        }

    }

    class EmpreendGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Empreend> map = [:]

        public EmpreendGenerator(Context context)
        {
            populateEmpreends(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Empreend(
                        cdEmpreend: key?.toString(),
                        nmEmpreend: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Empreendimentos"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_EMPREEND
        }

        @Override
        String getGroupBy()
        {
            return "b.cdEmpreend"
        }

        @Override
        String getGroupByIqa()
        {
            return "b.cdEmpreend"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdEmpreend"
        }

        @Override
        Object convertKey(Object key)
        {
            return key
        }

        void populateEmpreends(Context context)
        {
            List<Integer> cenarios = context.cenarios

            String select = ""

            if (cenarios.size() > 0)
                select = """
                    SELECT
                          a.cdEmpreend
                        , COALESCE(f.nmEmpreend, a.nmEmpreend)
                        , a.cdEmpreendProjeto
                        , a.cdRegiao
                        , a.cdRegional
                        , a.cdSubRegional
                        , COALESCE(g.nmEmpreend, b.nmEmpreend) AS nmEmpreendProjeto
                        , c.dsRegiao
                        , d.dsRegional
                        , e.dsSubRegional
                        , a.cdTipoEmpreend
                    FROM tb_CenarioOrcamentoEmpreend a
                    LEFT JOIN tb_CenarioOrcamentoEmpreend b
                      ON a.cdCenario = b.cdCenario
                      AND a.cdEmpreendProjeto = b.cdEmpreend
                    LEFT JOIN tb_Regiao c
                      ON a.cdRegiao = c.cdRegiao
                    LEFT JOIN tb_Regional d
                      ON a.cdRegional = d.cdRegional
                    LEFT JOIN tb_SubRegional e
                      ON a.cdSubRegional = e.cdSubRegional
                    LEFT JOIN tb_Empreend f 
                        ON f.cdEmpreend = a.cdEmpreend
                    LEFT JOIN tb_Empreend g
                        ON g.cdEmpreend = a.cdEmpreend
                    WHERE a.cdCenario IN (${cenarios.join(', ')})
                """
            else
                select = """
                    SELECT
                          a.cdEmpreend
                        , a.nmEmpreend
                        , a.cdEmpreendProjeto
                        , a.cdRegiao
                        , a.cdRegional
                        , a.cdSubRegional
                        , b.nmEmpreend AS nmEmpreendProjeto
                        , c.dsRegiao
                        , d.dsRegional
                        , e.dsSubRegional
                        , a.cdTipoEmpreend
                    FROM tb_empreend a
                    LEFT JOIN tb_Empreend b
                      ON a.cdEmpreendProjeto = b.cdEmpreend
                    LEFT JOIN tb_Regiao c
                      ON a.cdRegiao = c.cdRegiao
                    LEFT JOIN tb_Regional d
                      ON a.cdRegional = d.cdRegional
                    LEFT JOIN tb_SubRegional e
                      ON a.cdSubRegional = e.cdSubRegional
                """

            PreparedStatement ps = apr.connection.prepareStatement(select)

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Empreend vo = new Empreend(
                        cdEmpreend: rs.getString(1),
                        nmEmpreend: rs.getString(2),
                        cdEmpreendProjeto: rs.getString(3),
                        cdRegiao: rs.getInt(4),
                        cdRegional: rs.getInt(5),
                        cdSubRegional: rs.getInt(6),
                        nmEmpreendProjeto: rs.getString(7),
                        dsRegiao: rs.getString(8),
                        dsRegional: rs.getString(9),
                        dsSubRegional: rs.getString(10),
                        cdTipoEmpreend: rs.getInt(11),
                )

                map[vo.cdEmpreend] = vo
            }

            rs.close()
            ps.close()
        }

    }

    class ProjetoConsolidadoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Empreend> map = [:]

        public ProjetoConsolidadoGenerator(Context context)
        {
            populateEmpreends(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Empreend(
                        cdEmpreend: key?.toString(),
                        nmEmpreend: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Projetos Consolidados"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_PROJ_CONSOLIDADO
        }

        @Override
        String getGroupBy()
        {
            return "b.cdEmpreendProjeto"
        }

        @Override
        String getGroupByIqa()
        {
            return "b.cdEmpreendProjeto"
        }
        
        @Override
        String getGroupByVO()
        {
            //Nesse caso é necessário criar no mapa o racional do getGroupBy com um atributo que tenha o nome de cdEmpreendProjeto
            return "cdEmpreendProjeto"
        }

        @Override
        Object convertKey(Object key)
        {
            return key
        }

        void populateEmpreends(Context context)
        {
            List<Integer> cenarios = context.cenarios

            String select = ""
            if (cenarios.size() > 0) {
                select = """
                    SELECT 
                        a.cdEmpreend, 
                        COALESCE(b.nmEmpreend, a.nmEmpreend) 
                    FROM tb_CenarioOrcamentoEmpreend a
                    LEFT JOIN tb_Empreend b
                      ON a.cdEmpreendProjeto = b.cdEmpreend
                    WHERE a.cdCenario IN (${cenarios.join(', ')})
                """
            }
            else {
                select = """
                    SELECT 
                        cdEmpreend, 
                        nmEmpreend 
                    FROM tb_Empreend
                """
            }

            PreparedStatement ps = apr.connection.prepareStatement(select)

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Empreend vo = new Empreend(
                        cdEmpreend: rs.getString(1),
                        nmEmpreend: rs.getString(2)
                )

                map[vo.cdEmpreend] = vo
            }

            rs.close()
            ps.close()
        }

    }

    class MacroProjetoConsolidadoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Empreend> map = [:]

        public MacroProjetoConsolidadoGenerator(Context context)
        {
            populateEmpreends(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Empreend(
                        cdEmpreend: key?.toString(),
                        nmEmpreend: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Macro Projetos"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_MACRO_PROJ_CONSOLIDADO
        }

        @Override
        String getGroupBy()
        {
            return "b.cdEmpreendMacroProjeto"
        }

        @Override
        String getGroupByIqa()
        {
            return "b.cdEmpreendMacroProjeto"
        }
        
        @Override
        String getGroupByVO()
        {
            //Nesse caso é necessário criar no mapa o racional do getGroupBy com um atributo que tenha o nome de cdEmpreendMacroProjeto
            return "cdEmpreendMacroProjeto"
        }

        @Override
        Object convertKey(Object key)
        {
            return key
        }

        void populateEmpreends(Context context)
        {
            List<Integer> cenarios = context.cenarios

            String select = ""
            if (cenarios.size() > 0)
                select = "SELECT cdEmpreend, nmEmpreend FROM tb_CenarioOrcamentoEmpreend WHERE cdCenario IN (${cenarios.join(', ')})"
            else
                select = "SELECT cdEmpreend, nmEmpreend FROM tb_Empreend"

            PreparedStatement ps = apr.connection.prepareStatement(select)

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Empreend vo = new Empreend(
                        cdEmpreend: rs.getString(1),
                        nmEmpreend: rs.getString(2)
                )

                map[vo.cdEmpreend] = vo
            }

            rs.close()
            ps.close()
        }

    }

    static class EmpresaDivisaoNodeData implements NodeData
    {
        Object key
        String dsEmpreend
        String cdEmpresaERP
        String cdEmpreend
        int cdEmpresa

        @Override
        Object getKey()
        {
            return key
        }

        @Override
        String getDescr()
        {
            return dsEmpreend?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${keyToDisplay} ${dsEmpreend?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${cdEmpresaERP}/${cdEmpreend}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }

        boolean equals(o)
        {
            if (this.is(o)) return true
            if (getClass() != o.class) return false

            EmpresaDivisaoNodeData that = (EmpresaDivisaoNodeData) o

            if (key != that.key) return false

            return true
        }

        int hashCode()
        {
            return (key != null ? key.hashCode() : 0)
        }


    }

    class EmpresaDivisaoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, EmpresaDivisaoNodeData> map = [:]

        public EmpresaDivisaoGenerator(Context context)
        {
        }

        @Override
        NodeData getData(Object key)
        {
            EmpresaDivisaoNodeData vo = map[key]
            if (!vo) {
                vo = new EmpresaDivisaoNodeData()

                map[key] = vo

                vo.key = key

                def (String cdEmpresa, String cdEmpreend) = (key as String).split('/')

                vo.cdEmpresa = Util.iVal(cdEmpresa)
                vo.cdEmpreend = cdEmpreend

                PreparedStatement ps = apr.connection.prepareStatement("""
                        SELECT
                              b.dsEmpresa
                            , a.nmEmpreend
                            , b.cdEmpresaERP
                            FROM tb_empreend a, tb_empresa b
                            WHERE a.cdEmpreend = ?
                            AND b.cdEmpresa = ?
                    """)

                ps.setString(1, cdEmpreend)
                ps.setInt(2, vo.cdEmpresa)

                ResultSet rs = ps.executeQuery()

                if (rs.next())
                {
                    vo.dsEmpreend = rs.getString(1) + "/" + rs.getString(2)
                    vo.cdEmpresaERP = rs.getString(3)
                }
                else
                {
                    vo.dsEmpreend = vo.key as String
                    vo.cdEmpresaERP = vo.cdEmpresa as String
                }

                rs.close()
                ps.close()
            }

            return vo
        }

        @Override
        String getDescr()
        {
            return "Empresas e Divisões"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_EMPRESA_DIVISAO
        }

        @Override
        String getGroupBy()
        {
            return "CAST(a.cdEmpresa AS VARCHAR) || '/' || a.cdNucleo"
        }

        @Override
        String getGroupByIqa()
        {
            return "CAST(a.cdEmpresa AS VARCHAR) || '/' || a.cdNucleo"
        }
        
        @Override
        String getGroupByVO()
        {
            //Nesse caso é necessário criar no mapa o racional do getGroupBy com um atributo que tenha o nome de cdEmpresaEmpreend
            return "cdEmpresaEmpreend"
        }

        @Override
        Object convertKey(Object key)
        {
            return key
        }

    }

    static class Regiao implements NodeData
    {
        int cdRegiao
        String dsRegiao

        @Override
        Object getKey()
        {
            return cdRegiao
        }

        @Override
        String getDescr()
        {
            return dsRegiao?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdRegiao} ${dsRegiao?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }
    }

    class RegiaoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Regiao> map = [:]

        public RegiaoGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Regiao(
                        cdRegiao: key,
                        dsRegiao: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Regiões"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_REGIAO
        }

        @Override
        String getGroupBy()
        {
            return "b.cdRegiao"
        }

        @Override
        String getGroupByIqa()
        {
            return "b.cdRegiao"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdRegiao"
        }

        @Override
        Object convertKey(Object key)
        {
            if (key instanceof Integer)
                return key

            return Util.iVal(key)
        }
        
        @Override
        Collection<Object> sort(Collection<Object> collection) 
        {
            List<Object> sortedKeys = [1, 2, 3, 7, 4]
            int totalKeys = sortedKeys.size()
            
            List<Object> result = collection.sort({ n1, n2 -> n1.data.key <=> n2.data.key })
            
            for (int index = totalKeys; index >= 0; index--) {
                Object key = sortedKeys[index]
                
                Object obj = result.find({n -> n.data.key == key })
                
                if (obj != null) {
                    result.remove(obj)
                    result.add(0, obj)
                }
            }
            
            return result
        }

        void populate(Context context)
        {
            PreparedStatement ps = apr.connection.prepareStatement("SELECT cdRegiao, dsRegiao FROM tb_Regiao")

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Regiao vo = new Regiao(
                        cdRegiao: rs.getInt(1),
                        dsRegiao: rs.getString(2)
                )

                map[vo.cdRegiao] = vo
            }

            rs.close()
            ps.close()
        }

    }

    static class Regional implements NodeData
    {
        int cdRegional
        String dsRegional

        @Override
        Object getKey()
        {
            return cdRegional
        }

        @Override
        String getDescr()
        {
            return dsRegional?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdRegional} ${dsRegional?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }

    }

    class RegionalGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Regional> map = [:]

        public RegionalGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Regional(
                        cdRegional: key,
                        dsRegional: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Regionais"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_REGIONAL
        }

        @Override
        String getGroupBy()
        {
            return "b.cdRegional"
        }

        @Override
        String getGroupByIqa()
        {
            return "b.cdRegional"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdRegional"
        }

        @Override
        Object convertKey(Object key)
        {
            if (key instanceof Integer)
                return key

            return Integer.parseInt(key?.toString())
        }

        void populate(Context context)
        {
            PreparedStatement ps = apr.connection.prepareStatement("SELECT cdRegional, dsRegional FROM tb_Regional")

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Regional vo = new Regional(
                        cdRegional: rs.getInt(1),
                        dsRegional: rs.getString(2)
                )

                map[vo.cdRegional] = vo
            }

            rs.close()
            ps.close()
        }

    }

    static class SubRegional implements NodeData
    {
        int cdSubRegional
        String dsSubRegional

        @Override
        Object getKey()
        {
            return cdSubRegional
        }

        @Override
        String getDescr()
        {
            return dsSubRegional?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdSubRegional} ${dsSubRegional?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }

    }

    class SubRegionalGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, SubRegional> map = [:]

        public SubRegionalGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new SubRegional(
                        cdSubRegional: key,
                        dsSubRegional: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Sub-Regionais"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_SUBREGIONAL
        }

        @Override
        String getGroupBy()
        {
            return "b.cdSubRegional"
        }

        @Override
        String getGroupByIqa()
        {
            return "b.cdSubRegional"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdSubRegional"
        }

        @Override
        Object convertKey(Object key)
        {
            if (key instanceof Integer)
                return key

            return Integer.parseInt(key?.toString())
        }

        void populate(Context context)
        {
            PreparedStatement ps = apr.connection.prepareStatement("SELECT cdSubRegional, dsSubRegional FROM tb_SubRegional")

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                SubRegional vo = new SubRegional(
                        cdSubRegional: rs.getInt(1),
                        dsSubRegional: rs.getString(2)
                )

                map[vo.cdSubRegional] = vo
            }

            rs.close()
            ps.close()
        }

    }

    static class CentroCusto implements NodeData
    {
        int cdCentroCusto
        String dsCentroCusto

        @Override
        Object getKey()
        {
            return cdCentroCusto
        }

        @Override
        String getDescr()
        {
            return dsCentroCusto?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            return "${cdCentroCusto} ${dsCentroCusto?.toUpperCase()}"
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }

    }

    class CentroCustoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, CentroCusto> map = [:]

        public CentroCustoGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new CentroCusto(
                        cdCentroCusto: key,
                        dsCentroCusto: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Centros de Custo"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_CENTRO_CUSTO
        }

        @Override
        String getGroupBy()
        {
            return "COALESCE(cdp.cd_centro_custo_real, a.cd_centro_custo)"
        }

        @Override
        String getGroupByIqa()
        {
            return "COALESCE(cdp.cd_centro_custo_real, a.cd_centro_custo)"
        }
        
        @Override
        String getGroupByVO()
        {
            return "cdCentroCusto"
        }

        @Override
        Object convertKey(Object key)
        {
            if (key instanceof Integer)
                return key

            return Integer.parseInt(key?.toString())
        }

        void populate(Context context)
        {
            PreparedStatement ps = apr.connection.prepareStatement("SELECT cd_centro_custo, ds_centro_custo FROM tb_centro_custo")

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                CentroCusto vo = new CentroCusto(
                        cdCentroCusto: rs.getInt(1),
                        dsCentroCusto: rs.getString(2)
                )

                map[vo.cdCentroCusto] = vo
            }

            rs.close()
            ps.close()
        }

    }

    static class Departamento implements NodeData
    {
        String cdDepartamento
        String dsDepartamento

        @Override
        Object getKey()
        {
            return cdDepartamento
        }

        @Override
        String getDescr()
        {
            return dsDepartamento?.toUpperCase()
        }

        @Override
        Object getDisplay()
        {
            if (cdDepartamento == dsDepartamento)
                return dsDepartamento
            else
                return "${cdDepartamento} ${dsDepartamento?.toUpperCase()}".trim()
        }

        @Override
        String getKeyToDisplay()
        {
            return "${key}"
        }

        @Override
        void applyStyle(Cell cell)
        {
        }

    }

    class DepartamentoGenerator extends AbstractDataGenerator implements NodeDataGenerator
    {
        Map<Object, Departamento> map = [:]

        public DepartamentoGenerator(Context context)
        {
            populate(context)
        }

        @Override
        NodeData getData(Object key)
        {
            NodeData data = map[key]
            if (data == null)
            {
                data = new Departamento(
                        cdDepartamento: key?.toString(),
                        dsDepartamento: key?.toString()
                )
                map[key] = data
            }

            return data
        }

        @Override
        String getDescr()
        {
            return "Departamento"
        }

        @Override
        String getType()
        {
            return GENERATOR_TYPE_DEPARTAMENTO
        }

        @Override
        String getGroupBy()
        {
            return "COALESCE(depto.nm_depto, COALESCE(cdp.cd_centro_custo_real, a.cd_centro_custo)::varchar)"
        }

        @Override
        String getGroupByIqa()
        {
            return getGroupBy()
        }
        
        @Override
        String getGroupByVO()
        {
            return "departamento"
        }

        @Override
        Object convertKey(Object key)
        {
            return key
        }

        void populate(Context context)
        {
            String select = """
                    select distinct 
                        COALESCE(b.nm_depto, a.cd_centro_custo::varchar) as cd_depto, 
                        COALESCE(b.nm_depto, a.ds_centro_custo) as nm_depto 
                        from tb_centro_custo a 
                            left join tb_centro_custo_depto b 
                                on a.cd_centro_custo = b.cd_centro_custo
                """
            PreparedStatement ps = apr.connection.prepareStatement(select)

            ResultSet rs = ps.executeQuery()

            while (rs.next())
            {
                Departamento vo = new Departamento(
                        cdDepartamento: rs.getString(1),
                        dsDepartamento: rs.getString(2)
                )

                map[vo.cdDepartamento] = vo
            }

            rs.close()
            ps.close()
        }

    }

    abstract class AbstractDataGenerator implements NodeDataGenerator
    {
        List<Object> filters = []

        @Override
        void addFilter(Object value)
        {
            if (value != null)
                filters << convertKey(value)
        }

        @Override
        List<Object> getFilters()
        {
            return filters
        }
        
        @Override
        Collection<Object> sort(Collection<Object> collection) 
        {
            return collection.sort({ n1, n2 -> n1.data.key <=> n2.data.key })    
        }
    }

    static class Indice
    {
        Object key
        String descr
        String sheetName
    }

    static class Info
    {
        Object key
        String descr
        int cdAnoMesPrevisao
    }

    class Context
    {
        Object extraContext

        LogUtil logger

        boolean crosstabAgrupavel = false

        boolean abrirItens
        boolean abrirNiveis
        boolean drilldown
        boolean combinado

        boolean cdConsiderarStandby

        boolean contrapartida

        boolean formatoTabular
        boolean createDescrCol = true

        int qtFormatoTabular = 0
        int firstColFormatoTabular
        int indexCdFormatoTabular
        int indexDsFormatoTabular

        int[] rcCdGroup = [-1, -1]
        int[] rcDsGroup = [-1, -1]

        Info info

        int periodoDe
        int periodoAte
        int anoDe
        int anoAte

        def exibirFiscal
        def exibirSocietario
        def exibirGerencial

        def dados = [:]

        int idJob

        List<Integer> cenarios = []
        Map<Integer, Item> itens = [:]
        Map<Object, NodeData> crosstabs = [:]
        NodeDataGenerator crosstabGenerator
        List<NodeDataGenerator> crosstabGeneratorList = []
        Map<Integer, List<Integer>> tipoGrupoServicoMap = [:]
        Map<String, List<Integer>> linhaCashflowMap = [:]
        List<NodeDataGenerator> nodeDataGeneratorList = []
        List<NodeDataGenerator> filterGeneratorList = []
        List<String> contaList = []
        List<String> contaCrList = []
        boolean mensal = true
        boolean filtroAnual = false

        Map<String, NodeDataGenerator> nodeDataGeneratorMap = [:]

        CrosstabNode crosstabRoot = new CrosstabNode(new NodeDataGeneric(
                key: "TOTAL",
                descr: "TOTAL"
        ),
                "TOTAL"
        )

        int runningRow
        int runningCol
        int totalRows
        int indexSheetCashflow
        int colNodeDescr
        int colNodeKey

        Range rangeNodeRef

        Node root = new Node(new NodeDataGeneric(
                key: "TOTAL",
                descr: "T O T A L"
        ),
                "TOTAL"
        )

        List<Indice> indiceList = []

        int cdAnoMesBaseSOA

        int getCdAnoMesPrevisao() {
            return this.info.cdAnoMesPrevisao
        }

        void setCdAnoMesPrevisao(int cdAnoMes) {
            this.info.cdAnoMesPrevisao = cdAnoMes
        }

        void fillCrosstabNodeRange() {
            if (crosstabGenerator && crosstabGenerator instanceof FillNodeRange) {
                for (Object key : ((FillNodeRange)crosstabGenerator).keys) {
                    getCrosstab(key)
                }
            }

            if (crosstabGeneratorList) {
                Boolean [] fillNodes = new Boolean[crosstabGeneratorList.size()]

                for (int g = 0; g < crosstabGeneratorList.size(); g++) {
                    fillNodes[g] = crosstabGeneratorList[g] instanceof FillNodeRange
                }

                fillCrosstabNodeRangeRecursive crosstabRoot, fillNodes, 0
            }
        }

        void fillCrosstabNodeRangeRecursive(CrosstabNode crosstabNode, Boolean[] fillNodes, int indexChild)
        {
            if (crosstabNode.hasChildren() && fillNodes[indexChild])
            {
                for (Object key : ((FillNodeRange)crosstabGeneratorList[indexChild]).keys) {
                    CrosstabNode _crosstabNode = crosstabNode

                    boolean use = true

                    while (_crosstabNode.parent != null)
                    {
                        if ((_crosstabNode.data instanceof DateFilter && !_crosstabNode.data.isValidByAnoMes(crosstabGeneratorList[indexChild], key))) {
                            use = false
                            break
                        }

                        _crosstabNode = _crosstabNode.parent
                    }

                    if (use)
                        crosstabNode.getChild(key, crosstabGeneratorList[indexChild], true)
                }
            }

            for (CrosstabNode crosstabChild : crosstabNode.sortedChildren)
                fillCrosstabNodeRangeRecursive crosstabChild, fillNodes, indexChild + 1
        }

        NodeDataGenerator getNodeDataGenerator(String name)
        {
            //println "NODE : ${name}"

            NodeDataGenerator nodeDataGenerator = nodeDataGeneratorMap[name]
            if (nodeDataGenerator == null)
            {
                switch (RptHelper.tsu(name))
                {
                    case GENERATOR_TYPE_PROJ_CONSOLIDADO:
                        nodeDataGenerator = new ProjetoConsolidadoGenerator(this)
                        break
                    case GENERATOR_TYPE_MACRO_PROJ_CONSOLIDADO:
                        nodeDataGenerator = new MacroProjetoConsolidadoGenerator(this)
                        break
                    case GENERATOR_TYPE_EMPREEND:
                        nodeDataGenerator = new EmpreendGenerator(this)
                        break
                    case GENERATOR_TYPE_REGIAO:
                        nodeDataGenerator = new RegiaoGenerator(this)
                        break
                    case GENERATOR_TYPE_REGIONAL:
                        nodeDataGenerator = new RegionalGenerator(this)
                        break
                    case GENERATOR_TYPE_SUBREGIONAL:
                        nodeDataGenerator = new SubRegionalGenerator(this)
                        break
                    case GENERATOR_TYPE_EMPRESA:
                        nodeDataGenerator = new EmpresaGenerator(this)
                        break
                    case GENERATOR_TYPE_EMPRESA_DIVISAO:
                        nodeDataGenerator = new EmpresaDivisaoGenerator(this)
                        break
                    case GENERATOR_TYPE_CENTRO_CUSTO:
                        nodeDataGenerator = new CentroCustoGenerator(this)
                        break
                    case GENERATOR_TYPE_DEPARTAMENTO:
                        nodeDataGenerator = new DepartamentoGenerator(this)
                        break
                    case GENERATOR_TYPE_PERIODO:
                    case GENERATOR_TYPE_PERIODO_ERP:
                        nodeDataGenerator = new PeriodoGenerator(this)
                        break
                    case GENERATOR_TYPE_ANO:
                        nodeDataGenerator = new AnoGenerator(this)
                        break
                    case GENERATOR_TYPE_ESTUDO_ACOMP:
                        nodeDataGenerator = new EstudoAcompGenerator(this)
                        break
                    case GENERATOR_TYPE_NEGOCIO:
                        nodeDataGenerator = new NegocioGenerator(this)
                        break
                    case GENERATOR_TYPE_GRUPO_SERVICO:
                        nodeDataGenerator = new GrupoServicoGenerator(this)
                        break
                    default:
                        throw new Exception("Agrupamento '${name}' desconhecido!")
                }

                nodeDataGeneratorMap[name] = nodeDataGenerator
            }

            return nodeDataGenerator
        }

        void addFilter(String name, def values)
        {
            if (values)
            {
                NodeDataGenerator filterGenerator = getNodeDataGenerator(name)
                filterGeneratorList << filterGenerator

                for (Object value : values)
                    filterGenerator.addFilter(value)

            }
        }

        NodeData getCrosstab(Object crosstabKey)
        {
            NodeData crosstab = crosstabs[crosstabKey]
            if (crosstab == null)
            {
                crosstab = crosstabGenerator.getData(crosstabKey)
                crosstabs[crosstabKey] = crosstab
            }

            return crosstab
        }

        void addCrosstabNode(Object[] crosstabKeys) {

            CrosstabNode crosstabNode = this.crosstabRoot

            for (int i = 0; i < crosstabKeys.size(); i++) {
                crosstabNode = crosstabNode.getChild(crosstabKeys[i], this.crosstabGeneratorList[i], true)
            }
        }

        String getAgrupamentos()
        {
            String agrupamentos = ""
            for (NodeDataGenerator nodeDataGenerator : nodeDataGeneratorList)
                agrupamentos += ", " + nodeDataGenerator.descr

            if (agrupamentos.length() > 0)
                return agrupamentos.substring(2)
            else
                return "TOTAL"
        }

        String getFiltros()
        {
            String filter = ""

            for (NodeDataGenerator filterGenerator : filterGeneratorList)
            {
                if (filterGenerator.filters.size() > 0)
                {
                    if (filter.size() > 0)
                        filter += " | "
                    filter += filterGenerator.descr + ": "
                    int count = 0
                    for (Object value : filterGenerator.filters)
                    {
                        if (++count > 1)
                            filter += ', '

                        NodeData data = filterGenerator.getData(value)
                        filter += data.key + "-" + data.descr
                    }
                }
            }

            if (filter)
                filter = " e Filtro de " + filter

            return filter
        }

        void log(String text) {
            logger.log(text)
        }

        void erro(String text) {
            logger.erro(text)
        }

        boolean isSaldoInicial(int index) {
            Item item = itens[index]

            if (item?.linhasCashflow) {
                return item.linhasCashflow.containsValue(1)
            }

            return false
        }
    }

    static class NodeConta
    {
        String cdConta
        String dsConta
        int vlSinal
        String cdTag
        List<NodeConta> nodes = []
        NodeConta parent
    }

    interface NodeData
    {
        Object getKey()
        String getDescr()
        Object getDisplay()
        String getKeyToDisplay()
        void applyStyle(Cell cell)
    }

    interface NodeDataGenerator
    {
        String getDescr()
        NodeData getData(Object key)
        String getType()
        String getGroupBy()
        String getGroupByIqa()
        String getGroupByVO()
        Object convertKey(Object key)
        void addFilter(Object value)
        List<Object> getFilters()
        Collection<Object> sort(Collection<Object> collection)
    }

    static class Condition
    {
        boolean menor = false
        boolean igual = false
        boolean maior = false

        public Condition(String cond)
        {
            compile(cond)
        }

        private void compile(String cond)
        {
            if (cond)
                if (cond == '<')
                    menor = true
                else if (cond == '<=')
                    menor = igual = true
                else if (cond == '>')
                    maior = true
                else if (cond == '>=')
                    maior = igual = true
                else if (cond == '=')
                    igual = true
                else if (cond == '<>')
                    menor = maior = true
        }

        public boolean validate(Object o1, Object o2)
        {
            if (menor)
                if (o1 < o2)
                    return true

            if (igual)
                if (o1 == o2)
                    return true

            if (maior)
                if (o1 > o2)
                    return true

            return false
        }

    }
}                     