package com.main;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {
	private static Workbook wb;
	private static Sheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;
	private static Row row;
	private static Cell cell;

	public static void main(String[] args) throws Exception {

		fis = new FileInputStream("./testExcel.xlsx");
		wb = WorkbookFactory.create(fis);
		sh = wb.getSheet("Sheet1");
		int noOfRows = sh.getLastRowNum();
		System.out.println(noOfRows);

		row = sh.createRow(1);
		cell = row.createCell(0);

		cell.setCellValue("Sidharth Das");

		cell = row.createCell(1);
		cell.setCellValue("password123");

		cell = row.createCell(2);
		cell.setCellValue("Fuckyou");
		for(int i = 3; i <10; i++) {
			cell = row.createCell(i);
			cell.setCellValue("Fuckyou");
		}

		System.out.println(cell.getStringCellValue());

		fos = new FileOutputStream("./testExcel.xlsx");
		wb.write(fos);
		fos.flush();

	}

}
