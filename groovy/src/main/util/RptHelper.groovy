package src.main.util

import java.sql.Array
import java.sql.Connection

class RptHelper
{
    static String tsu(String text) {
        return Util.noAccent(text)?.trim()?.toUpperCase()
    }
}        