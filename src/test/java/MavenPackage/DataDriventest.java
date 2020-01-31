package MavenPackage;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriventest {

	public static void main() throws IOException{
		
		FileInputStream fp= new FileInputStream("C://TestData//DemoData.xlsx");
        
		@SuppressWarnings("resource")
		XSSFWorkbook wbook= new XSSFWorkbook(fp);
		
		int sheets= wbook.getNumberOfSheets();
		
		for (int i=0; i<sheets;i++)
		{
			if(wbook.getSheetName(i).equalsIgnoreCase("DemoData"))
			{
				XSSFSheet sheet= wbook.getSheetAt(0);
				
				Iterator<Row> rows=sheet.iterator();
				Row firstrow= rows.next();
				
				Iterator<Cell> cel= firstrow.cellIterator();
				int k=0;
				
				int column=0;
				while(cel.hasNext())
				{
					Cell value=cel.next();
					
					if(value.getStringCellValue().equals("Testcases"));
					{
						column=k;
					}
					k++;

					System.out.println("column count si:"+column);
				}
			}
			
		}
			
		
	}

}
