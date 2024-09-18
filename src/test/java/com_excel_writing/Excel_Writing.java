package com_excel_writing;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Writing {
	
	public static void main(String[] args) throws IOException {

		// it converts into writing mode
		FileOutputStream fos=new FileOutputStream("F:\\Eclipse work space\\MavenTeam2930\\ExcelFileDocument\\ExcelWriting22.xlsx");
		
		XSSFWorkbook wb=new XSSFWorkbook(); // it is XML spread sheet format
	    XSSFSheet sheet=wb.createSheet(); // it will represents sheets
		
		
	    Scanner sc=new Scanner(System.in);
	    
	    for(int r=0; r<=3; r++) { //it will represents rows -->4
	    	
	    //Create row
	    	
	    		XSSFRow row=sheet.createRow(r);
	    		
	    for(int c=0; c<=2; c++ ) { //it will represents columns-->3
	    	
	    	System.out.println("Enter values");
	    	
	    	String values=sc.next(); // it will accepte string related values
	    	
	    	row.createCell(c).setCellValue(values);	  
	    	
	    }	
	    }
	    
	    wb.write(fos); //writing
	    wb.close(); //close-XSSFWorkbook
	    fos.close();//close-filepath
	
	    System.out.println("Values entering is done");
	}
	

}
