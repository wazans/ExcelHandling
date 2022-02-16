package ExcelHandling;

import java.io.FileInputStream;

import java.io.File;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Topic1_readSpecificCellValue {

	public static void main(String[] args) throws IOException {
		
		//https://www.toolsqa.com/selenium-webdriver/excel-in-selenium/
		//Create an object of File class to open xlsx file
        File file =new File("TestData.xlsx");
        
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        
        //Creating workbook instance that refers to .xls file
        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
        
        //Creating a Sheet object using the sheet Name
        XSSFSheet sheet=wb.getSheet("STUDENT_DATA");
        
        //Create a row object to retrieve row at index 1
        XSSFRow row2=sheet.getRow(2);
        
        //Create a cell object to retreive cell at index 3
        XSSFCell cell=row2.getCell(4);
        
        //Get the address in a variable
        String adress= cell.getStringCellValue();
        
        //Printing the address
        System.out.println("adress is :"+ adress);

	}

}
