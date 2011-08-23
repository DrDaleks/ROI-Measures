package plugins.adufour.roi;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;

import jxl.Cell;
import jxl.Sheet;
import jxl.write.WritableWorkbook;

public class ExportCSV
{
	public static void export(WritableWorkbook workbook, File file, boolean skipHiddenCells) throws IOException
	{
		try
		{
			OutputStream out = new FileOutputStream(file);
			OutputStreamWriter osw = new OutputStreamWriter(out, "UTF8");
			BufferedWriter bw = new BufferedWriter(osw);
			
			for (int sheet = 0; sheet < workbook.getNumberOfSheets(); sheet++)
			{
				Sheet s = workbook.getSheet(sheet);
				
				if ((skipHiddenCells) && (s.getSettings().isHidden()))
					continue;
				bw.write("*** " + s.getName() + " ****");
				bw.newLine();
				
				Cell[] row = null;
				
				for (int i = 0; i < s.getRows(); i++)
				{
					row = s.getRow(i);
					
					if (row.length > 0)
					{
						if ((!skipHiddenCells) || (!row[0].isHidden()))
						{
							bw.write(row[0].getContents());
						}
						
						for (int j = 1; j < row.length; j++)
						{
							bw.write(44);
							if ((skipHiddenCells) && (row[j].isHidden()))
								continue;
							bw.write(row[j].getContents());
						}
						
					}
					
					bw.newLine();
				}
			}
			
			bw.flush();
			bw.close();
		}
		catch (UnsupportedEncodingException e)
		{
			System.err.println(e.toString());
		}
	}
}
