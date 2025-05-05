import java.sql.Statement
import java.sql.ResultSet
import java.sql.Connection
import java.sql.DriverManager

import com.aspose.cells.*

void criandoConexaoSql() {

    println "Buscando variáveis de ambiente."

    String HOST = System.getProperty("HOST")
    String DATABASE = System.getProperty("DATABASE")
    String USER = System.getProperty("USER")
    String PASSWORD = System.getProperty("PASSWORD")

    println "Conectando com o banco de dados." 

    String url = "jdbc:postgresql://${HOST}/${DATABASE}";
    Properties props = new Properties();
    props.setProperty("user", USER);
    props.setProperty("password", PASSWORD);
    Connection conn = DriverManager.getConnection(url, props);

    Statement st = conn.createStatement();
    ResultSet rs = st.executeQuery("SELECT count(*) FROM tb_ev_unidade_erp WHERE cd_ano_mes_base = 202501");

    while (rs.next()) {
        println "RESULTADO: ${rs.getInt(1)}"
    }

    rs.close();
    st.close();
    
}

criandoConexaoSql()
Workbook workbook = null