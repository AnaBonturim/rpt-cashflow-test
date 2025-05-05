package src.main

import java.sql.Statement
import java.sql.ResultSet
import java.sql.Connection
import java.sql.DriverManager

import groovy.sql.*

public class Apr {

    static Connection connection = null

    public Apr() {

    }

    void throwException(String text) {
        throw Exception(text)
    }

    Connection getConnection() {

        if (!this.connection) {
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
            
            this.connection = DriverManager.getConnection(url, props);
        }

        return this.connection
    }

    byte[] getByData() {

        Sql sql = new Sql(this.getConnection())

        String select = """
            SELECT
                by_data AS "byData"
            FROM tb_template
            WHERE cd_tag = 'PE-CASHFLOW'
        """

        def row = sql.firstRow(select)

        return row.byData
    }
}