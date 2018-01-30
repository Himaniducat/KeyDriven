package excelExportAndFileIO;

import java.io.File;
import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFile {

	public Sheet readExcel(String filePath , String fileName , String sheetName) throws IOException
	{  //create an object of file class to open xlsx file
		System.out.println(filePath);
		File file = new File(filePath + "\\src\\" + fileName);
		
		//create an object of fileInputStream to read excel file
		FileInputStream inputstream = new FileInputStream(file);
		
		Workbook workbook = null;
		
		//find the file extension by using substring method of string
	  	String fileExtensionName = fileName.substring(fileName.indexOf("."));
	  	
	    if(fileExtensionName.equals(".xlsx"))
	    {
	    	workbook  = new XSSFWorkbook(inputstream);
	    	
	    }
	    else if(fileExtensionName.equals(".xls"))
	    {
	    	workbook = new HSSFWorkbook(inputstream);
	    }
	    
	    Sheet sheet = workbook.getSheet(sheetName);
	    return sheet;
	}

}
