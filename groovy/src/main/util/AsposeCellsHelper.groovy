package src.main.util

import com.aspose.cells.DateTime;
import java.util.GregorianCalendar;

class AsposeCellsHelper {
    static DateTime amToDateTime(int cdAnoMes)
	{
		if (cdAnoMes == 0)
			return null;

		return new DateTime(new GregorianCalendar((int)(cdAnoMes / 100), cdAnoMes % 100 - 1, 1).getTime());
	}
}