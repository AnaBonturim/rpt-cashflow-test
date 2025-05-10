package src.main

class Dados {

    int cdCenario = cdCenario
    List<String> agrupamentoList = []
    int cdConsiderarStandby = cdConsiderarStandby
    boolean formatoTabular

    Dados(int cdCenario, List<String> agrupamentoList, int cdConsiderarStandby) {
        this.cdCenario = cdCenario
        this.agrupamentoList = agrupamentoList
        this.cdConsiderarStandby = cdConsiderarStandby
    }
}


Apr apr = new Apr()
RptCashflow report = new RptCashflow(apr)
Dados dados = new Dados(39822, ["regional", "empreend", "periodo"], 1)
dados.formatoTabular = true

println("Começando processo às ${new Date().format("HH:mm:ss")}")

report.execute(dados)    
    
println("Terminando processo às ${new Date().format("HH:mm:ss")}")