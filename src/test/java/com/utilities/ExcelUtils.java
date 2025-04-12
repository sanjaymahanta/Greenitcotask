package com.utilities;



import java.io.FileInputStream;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static XSSFWorkbook wb;
	public static XSSFSheet ws;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static CellStyle style;

	public static int getRowCount(String xlfile, String xlsheet) throws Exception {
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		int rowCount = ws.getLastRowNum();
		wb.close();
		fi.close();
		return rowCount;

	}

	public static int getCellCount(String xlfile, String xlsheet, int rownum) throws Exception {
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		int cellCount = row.getLastCellNum();
		wb.close();
		fi.close();
		return cellCount;

	}                         

	public static String getCellData(String xlfile, String xlsheet, int rownum, int colnum) throws Exception {
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(colnum);

		String data;
		try {
			// data = cell.toString();
			DataFormatter formatter = new DataFormatter();
			data = formatter.formatCellValue(cell);

		} catch (Exception e) {
			data = "";
		}

		wb.close();
		fi.close();
		return data;

	}
	
	public static void setCellData(String path, String sheetName, int rowNum, int colNum, String data) throws Exception {
	    FileInputStream fis = new FileInputStream(path);
	    Workbook workbook = WorkbookFactory.create(fis);
	    Sheet sheet = workbook.getSheet(sheetName);

	    if (sheet.getRow(rowNum) == null) {
	        sheet.createRow(rowNum);
	    }
	    if (sheet.getRow(rowNum).getCell(colNum) == null) {
	        sheet.getRow(rowNum).createCell(colNum);
	    }

	    sheet.getRow(rowNum).getCell(colNum).setCellValue(data);
	    
	    fis.close();  // Close input stream before writing

	    FileOutputStream fos = new FileOutputStream(path);
	    workbook.write(fos);

	    fos.close();  // Close output stream after writing
	    workbook.close();  // Close workbook to free memory
	}

}