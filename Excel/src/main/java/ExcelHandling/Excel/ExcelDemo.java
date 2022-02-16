package ExcelHandling.Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelDemo {

	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws FileNotFoundException {
		FileInputStream fis=new FileInputStream("D:\\checklist.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook();
	
		
		int sheets=wb.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{		if(wb.getSheetName(i).equalsIgnoreCase("Sheet3"))
			{
			
			XSSFSheet  sheet=wb.getSheetAt(i);
			
			Iterator<Row> row=sheet.iterator();
			Row firstRow=row.next();
			
			Iterator<Cell> cell=firstRow.cellIterator();
			int k=0;
			int column=0;
			while(cell.hasNext())
			{
				Cell value=cell.next();
				if(value.getStringCellValue().equalsIgnoreCase("Sheet3"))
				{
					column=k;
				}
				k++;
				
			}
			System.out.println(column);
		
			}
		}
		
		
		
		

}
}

