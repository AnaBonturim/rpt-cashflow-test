package src.main.util

public class LogUtil
{

    void log(String text)
    {
        println "${new Date().format("HH:mm:ss")} LOG: ${text}"
    }

    void logs(List<String> values)
    {
        if (!values)
            return

        for (String txt : values)
            log(txt)
    }

    void warning(String text)
    {
        println "${new Date().format("HH:mm:ss")} WARNING: ${text}"
    }

    void warnings(List<String> values)
    {
        if (!values)
            return

        for (String txt : values)
            warning(txt)
    }

    void erro(String text)
    {
        println "${new Date().format("HH:mm:ss")} ERRO: ${text}"
    }

    void erros(List<String> values)
    {
        if (!values)
            return

        for (String txt : values)
            erro(txt)
    }
}   