package testcases;

//import java.io.IOException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
//import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import excelExportAndFileIO.ExcelFile;
import operation.ReadObject;
import operation.UIOperation;

public class executeTestCase {
	
	@Test
	public void testLogin() throws Exception
	{
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\dell\\Desktop\\chromedriver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		ExcelFile file = new ExcelFile();
		ReadObject object = new ReadObject();
		Properties allObjects = object.getObjectRepository();
		UIOperation operation = new UIOperation(driver);
		//Read keyword sheet
		Sheet sheet = file.readExcel(System.getProperty("user.dir"), "Keydriven.xlsx", "Keyword Framework");
		
		//Find number of rows in excel file
		int rowcount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		
		//Create a loop over all the rows of excel file to read it
		for(int i = 0 ; i <= rowcount; i++)
		{
			//Loop over all the rows
			Row row = sheet.getRow(i);
			
			//Check if the first cell contain a value, if yes, That means it is the new testcase name
			if(this.check(row.getCell(0)).length() == 0)
			{
				//Print testcase detail on console
				System.out.println(this.check(row.getCell(1)) +"----"+ this.check(row.getCell(2)) +"-----"+ this.check(row.getCell(3))+"-----"+this.check(row.getCell(4)));
				
				//Call perform function to perform operation on UI
				operation.perform(allObjects, this.check(row.getCell(1)), this.check(row.getCell(2)), this.check(row.getCell(3)), this.check(row.getCell(4)));
			}
			
			else
			{
				System.out.println("New TestCase ->" + this.check(row.getCell(0)) + "Started");
			}
		}
		
		
	}
	
	
	 public String check(Cell c)
	 {
		return (c == null) ? "" : c.toString();
	 }
	

}
