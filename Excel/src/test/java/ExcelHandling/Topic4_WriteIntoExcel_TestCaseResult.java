package ExcelHandling;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Topic4_WriteIntoExcel_TestCaseResult {
	public static void main(String args[]) throws IOException, InterruptedException {
		// set the ChromeDriver path
		System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

		// Create an object of File class to open xls file

		File file = new File("Registration.xlsx");

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);

		// creating workbook instance that refers to .xls file
		XSSFWorkbook wb = new XSSFWorkbook(inputStream);

		// creating a Sheet object
		XSSFSheet sheet = wb.getSheet("STUDENT_DATA");

		// get all rows in the sheet
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		// Creating an object of ChromeDriver
		WebDriver driver = new ChromeDriver();

		// Navigate to the URL
		driver.get("https://demoqa.com/automation-practice-form");

		// Identify the WebElements for the student registration form
		WebElement firstName = driver.findElement(By.id("firstName"));
		WebElement lastName = driver.findElement(By.id("lastName"));
		WebElement email = driver.findElement(By.id("userEmail"));
		WebElement genderMale = driver.findElement(By.id("gender-radio-1"));
		WebElement mobile = driver.findElement(By.id("userNumber"));
		WebElement address = driver.findElement(By.id("currentAddress"));
		WebElement submitBtn = driver.findElement(By.id("submit"));

		// iterate over all the rows in Excel and put data in the form.
		for (int i = 1; i <= rowCount; i++) {
			//Convert all the values to "String"
        	DataFormatter formatter = new DataFormatter();
        	formatter.formatCellValue(sheet.getRow(i).getCell(4));
            //Enter the values read from Excel in firstname,lastname,mobile,email,address
            firstName.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
            lastName.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
            email.sendKeys(sheet.getRow(i).getCell(2).getStringCellValue());
            //Click on the gender radio button using javascript
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("arguments[0].click();", genderMale);
            mobile.sendKeys("9641782691");
         
            address.sendKeys(sheet.getRow(i).getCell(5).getStringCellValue());
            
            Thread.sleep(5000);
            
            //Click on submit button
            //submitBtn.click();
            //Used java Script executor to solve element not interactable exception
            JavascriptExecutor js1 = (JavascriptExecutor) driver;
            js1.executeScript("arguments[0].click();", submitBtn);
			// Verify the confirmation message
			WebElement confirmationMessage = driver
					.findElement(By.xpath("//div[text()='Thanks for submitting the form']"));

			// create a new cell in the row at index 6
			XSSFCell cell = sheet.getRow(i).createCell(6);

			// check if confirmation message is displayed
			if (confirmationMessage.isDisplayed()) {
				// if the message is displayed , write PASS in the excel sheet
				cell.setCellValue("PASS");

			} else {
				// if the message is not displayed , write FAIL in the excel sheet
				cell.setCellValue("FAIL");
			}

			// Write the data back in the Excel file
			FileOutputStream outputStream = new FileOutputStream("TestResult.xls");
			wb.write(outputStream);

			// close the confirmation popup
			 WebElement closebtn = driver.findElement(By.id("closeLargeModal"));
	            JavascriptExecutor js2 = (JavascriptExecutor) driver;
	            js2.executeScript("arguments[0].click();", closebtn);
	            //closebtn.click();
	            Thread.sleep(5000);

			// wait for page to come back to registration page after close button is clicked
			driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
		}

		// Close the workbook
		wb.close();

		// Quit the driver
		driver.close();
	}
}