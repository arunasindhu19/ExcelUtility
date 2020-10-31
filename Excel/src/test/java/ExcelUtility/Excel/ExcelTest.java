package ExcelUtility.Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTest {
	
	public ArrayList<String> getData(String Names) throws IOException
	{
		String path = System.getProperty("user.dir");
		String FilePath = path + "\\TestData.xlsx";
		FileInputStream fis = new FileInputStream(FilePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		ArrayList<String> details = new ArrayList<String>();
		
		int sheets = workbook.getNumberOfSheets();
		
		for(int i=0;i<sheets;i++)
		{
			
			XSSFSheet sheetName = workbook.getSheetAt(i);
			if(sheetName.getSheetName().equalsIgnoreCase("Credentials"))
			{
				Iterator<Row> rows = sheetName.rowIterator();
				Row firstRow = rows.next();
				
				Iterator<Cell> cells = firstRow.cellIterator();
				
				int a =0;
				int columnNum = 0;
				while(cells.hasNext())
				{
					Cell c = cells.next();
				    if(c.getStringCellValue().equalsIgnoreCase("Names"))
				    {
				    	columnNum = a;
				    }
				    a++;
				}
				
				while(rows.hasNext())
				{
					Row r = rows.next();
					 
					if(r.getCell(columnNum).getStringCellValue().equalsIgnoreCase(Names))
					{
						Iterator<Cell> ce = r.cellIterator();
						while(ce.hasNext())
						{
							details.add(ce.next().getStringCellValue());
							
						}
					}		 
					
				}
				 
			}
			
		}
		return details;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
	}

}
