package src.main.util

import java.io.InputStream;
import java.net.InetAddress;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.regex.Pattern;

class Util {

    static int addIndexToPeriodo(int cdAnoMesPrevisao, int qtPeriodos)
	{
		if (qtPeriodos == 0)
			return cdAnoMesPrevisao;

		int ano = cdAnoMesPrevisao / 100;
		int mes = cdAnoMesPrevisao % 100 + qtPeriodos;

		if (mes > 12)
		{
			while (mes > 12)
			{
				mes -= 12;
				++ano;
			}
		}
		else if (mes <= 0)
			while (mes <= 0)
			{
				mes += 12;
				--ano;
			}

		return ano * 100 + mes;
	}
    
    static final int iVal(String v)
	{
		if (v == null)
			return 0;

		DecimalFormat decfmt = new DecimalFormat();

		int r;

		try
		{
			r = decfmt.parse(v).intValue();
		}
		catch(ParseException e)
		{
			r = 0;
		}

		return r;
	}

	static final int iVal(Object v)
	{
		if (v == null)
			return 0;

		if (v instanceof Integer)
			return ((Integer)v).intValue();
		else if (v instanceof Number)
			return ((Number)v).intValue();
		else
			return iVal(v.toString());
	}

    static final String noAccent(String ptxt)
	{
		StringBuffer sbna = new StringBuffer();

		if (ptxt == null)
            return null;

		for (int i = 0; i < ptxt.length(); ++i)
		{
			char c = ptxt.charAt(i);

			switch(c)
			{
				case 'Ã':
					sbna.append('A');
					break;
				case 'À':
					sbna.append('A');
					break;
				case 'Á':
					sbna.append('A');
					break;
				case 'Â':
					sbna.append('A');
					break;
				case 'Ä':
					sbna.append('A');
					break;
				case 'Å':
					sbna.append('A');
					break;
				case 'à':
					sbna.append('a');
					break;
				case 'á':
					sbna.append('a');
					break;
				case 'â':
					sbna.append('a');
					break;
				case 'ã':
					sbna.append('a');
					break;
				case 'ä':
					sbna.append('a');
					break;
				case 'å':
					sbna.append('a');
					break;
				case 'Ç':
					sbna.append('C');
					break;
				case 'ç':
					sbna.append('c');
					break;
				case 'È':
					sbna.append('E');
					break;
				case 'É':
					sbna.append('E');
					break;
				case 'Ê':
					sbna.append('E');
					break;
				case 'Ë':
					sbna.append('E');
					break;
				case 'è':
					sbna.append('e');
					break;
				case 'é':
					sbna.append('e');
					break;
				case 'ê':
					sbna.append('e');
					break;
				case 'ë':
					sbna.append('e');
					break;
				case 'Ì':
					sbna.append('I');
					break;
				case 'Í':
					sbna.append('I');
					break;
				case 'Î':
					sbna.append('I');
					break;
				case 'Ï':
					sbna.append('I');
					break;
				case 'ì':
					sbna.append('i');
					break;
				case 'í':
					sbna.append('i');
					break;
				case 'î':
					sbna.append('i');
					break;
				case 'ï':
					sbna.append('i');
					break;
				case 'Ñ':
					sbna.append('N');
					break;
				case 'ñ':
					sbna.append('n');
					break;
				case 'Ò':
					sbna.append('O');
					break;
				case 'Ó':
					sbna.append('O');
					break;
				case 'Ô':
					sbna.append('O');
					break;
				case 'Õ':
					sbna.append('O');
					break;
				case 'Ö':
					sbna.append('O');
					break;
				case 'ò':
					sbna.append('o');
					break;
				case 'ó':
					sbna.append('o');
					break;
				case 'ô':
					sbna.append('o');
					break;
				case 'õ':
					sbna.append('o');
					break;
				case 'ö':
					sbna.append('o');
					break;
				case 'Ù':
					sbna.append('U');
					break;
				case 'Ú':
					sbna.append('U');
					break;
				case 'Û':
					sbna.append('U');
					break;
				case 'Ü':
					sbna.append('U');
					break;
				case 'ù':
					sbna.append('u');
					break;
				case 'ú':
					sbna.append('u');
					break;
				case 'û':
					sbna.append('u');
					break;
				case 'ü':
					sbna.append('u');
					break;
				case 'Ý':
					sbna.append('Y');
					break;
				case 'ý':
					sbna.append('Y');
					break;
				case 'ÿ':
					sbna.append('y');
					break;
				case '\'':
					sbna.append('.');
					break;

				default:
					sbna.append(c);
			}
		}

		return sbna.toString();
	}
}