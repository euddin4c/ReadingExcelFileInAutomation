package com.osa.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFileDemo{  
	public static void main(String[] args) throws IOException{
		for(int i=2;i<8;i++) {
		String user=readexcel("C:\\Users\\sdsaj\\OneDrive\\Desktop\\student.xls.xlsx","A"+i);
		String pass=readexcel("C:\\Users\\sdsaj\\OneDrive\\Desktop\\student.xls.xlsx","B"+i);
		System.out.println(user+"      "+pass);
		
	}
	}	
	
	
	public static String readexcel(String path, String CellAddress) throws IOException {
		String value=null;
		DataFormatter df=new DataFormatter();
		try {
		File file=new File(path);
		FileInputStream fi=new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(fi);
		XSSFSheet sheet=wb.getSheet("Sheet1");
		CellAddress ca=new CellAddress(CellAddress);
		Row row=sheet.getRow(ca.getRow());
		Cell cell=row.getCell(ca.getColumn());
		value=df.formatCellValue(cell).toString();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return value;
		
		
		
		
		
	}	
		
		
		
		
		
		
		
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	



