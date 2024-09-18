package com_excel_reading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Excel_Reading1 {
	
	public static void main(String[] args) throws IOException {
			
	//it will convert into reading mode
	FileInputStream fis=new FileInputStream("F:\\Eclipse work space\\MavenTeam2930\\ExcelFileDocument\\ExcelReading.xlsx");
		
		XSSFWorkbook wb=new XSSFWorkbook(fis); // it is XML spread sheet format
	    XSSFSheet sheet=wb.getSheet("Sheet1"); // it will represents sheets
	    
	    //identify rows and columns
	    
	    int rows=sheet.getLastRowNum();
	    int cols=sheet.getRow(1).getLastCellNum();
	    
	    
	    for(int i=0; i<=rows; i++) {// it will represent rows //0,1,2,3
	    	
	    	XSSFRow crow=sheet.getRow(i);
	    	
	    	for(int c=0; c<cols; c++) { //it will represents columns//0,1,2
	    		
	    		String values =crow.getCell(c).toString();
	    		
	    	System.out.print(values+ "   ");
	    	}
	    	System.out.println();
	    }
	
	}

}
