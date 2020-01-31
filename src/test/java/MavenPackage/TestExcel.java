package MavenPackage;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class TestExcel {

@Test
public void testExcel() throws IOException{

ArrayList<String> ar= new ArrayList<String>();
FileInputStream fp= new FileInputStream("C://TestData//DemoData.xlsx");
        
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

				}
				
				while(rows.hasNext())
				{
					Row r=rows.next();

					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase"));
					{
					Iterator<Cell> cv= r.cellIterator();
							while(cv.hasNext())
							{
								System.out.println(cv.next().getStringCellValue());
							}
								
					}
				}	
			}
			
		}
			
		
	}

}


