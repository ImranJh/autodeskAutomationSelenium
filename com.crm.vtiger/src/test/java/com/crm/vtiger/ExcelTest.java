package com.crm.vtiger;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelTest {
	
	public static void main(String[] args) throws Throwable {
		FileInputStream f= new FileInputStream("C:\\Users\\Hp\\Desktop\\Test Yentra Study Material\\Java.xlsx");
		Workbook wb = WorkbookFactory.create(f);
		Sheet sh = wb.getSheet("Imran");
		Row r=sh.getRow(2);
//		Cell c=r.getCell(1);
//		String data=c.getStringCellValue();
		Cell c=r.createCell(4);
		c.setCellValue("India");
		
		FileOutputStream fo = new FileOutputStream("C:\\Users\\Hp\\Desktop\\Test Yentra Study Material\\Java.xlsx");
		wb.write(fo);
		
		//System.out.println(data);
	}

}
