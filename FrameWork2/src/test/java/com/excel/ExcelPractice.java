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

public class ExcelPractice {

	public static void main(String[] args) throws IOException {
		
	//get into file
		
		File file = new File("D:\\Software Testing\\FrameWork2\\Excel");
		
		FileInputStream stream =  new FileInputStream(file);
		
		Workbook book =  new XSSFWorkbook(stream);
		
		Sheet sheet = book.getSheet("sheet1");
		
		Row row = sheet.getRow(1);
		
		Cell cell = row.getCell(1);
		
		System.out.println(cell);
		
		
		

		
		
		
		
		
		
		
		
		
		
		
		
		
		
	
		
	}
}
