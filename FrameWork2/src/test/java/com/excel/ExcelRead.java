package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		
		File file = new File("D:\\Software Testing\\FrameWork2\\Excel\\framework1.xlsx");
		
		//get into file
		//class name fileinputstream
		FileInputStream stream = new FileInputStream(file);
		
		//et into workbook
		//interface - workbook, classes - xssfworkbook
		Workbook book = new XSSFWorkbook(stream);
		
		//get into sheet, method called getsheet
		Sheet sheet = book.getSheet("sheet1");
		
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		
		System.out.println("Row  count is : " + physicalNumberOfRows);
		
		Row row = sheet.getRow(1);

		int physicalNumberOfCells = row.getPhysicalNumberOfCells();
		
		System.out.println("Cell  count is : " + physicalNumberOfCells);

		for(int i = 0 ; i<sheet.getPhysicalNumberOfRows(); i++){
		
		Row row2 = sheet.getRow(i);
		
		for (int j = 0; j<row2.getPhysicalNumberOfCells(); j++){
			
			Cell cell = row2.getCell(j);
			
			System.out.println(cell);
		}
		
		}
		
	}	
}
