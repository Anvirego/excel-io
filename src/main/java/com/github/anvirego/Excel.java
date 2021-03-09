package com.github.anvirego;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.anvirego.interfaces.ExcelInterface;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 3.0 03/2021.
 * Excel: Reads data from an Excel File. 
 */
public final class Excel implements ExcelInterface {
	private static Excel excel = null;
	private String excelFileName;
	private String excelSheetName;
	private String data;
	private Sheet sheetBook;
	private Row row;
	private String cellData;
	private double cellNumberData;
	
	public Excel(String excelFileName, String excelSheetName) throws FileNotFoundException, IOException {
		this.excelFileName = excelFileName;
		this.excelSheetName = excelSheetName;
		readDataExcel();
	}//Constructor
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public static Excel getInstance(String excelFileName, String excelSheetName) throws FileNotFoundException, IOException {
		System.out.println("==== Get Excel Instance =====");
		if(excel == null) {
			System.out.println("New Instance");
			excel = new Excel(excelFileName, excelSheetName);
			return excel;
		} else {
			System.out.println("Old Instance");
			return excel;
		}
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	private void readDataExcel() throws FileNotFoundException, IOException {
		System.out.println("::::: readDataExcel("+excelSheetName+") :::::");
		File file = new File(excelFileName);
		//Create an object of FileInputStream to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		if(inputStream != null) {
			Workbook excelWorkBook = null;
			//Find the file extension by splitting excelFileName in substrings and getting only extension
			if((excelFileName.substring(excelFileName.indexOf("."))).equals(".xlsx")) {
				excelWorkBook = new XSSFWorkbook(inputStream);
			} else if((excelFileName.substring(excelFileName.indexOf("."))).equals(".xls")) {
				excelWorkBook = new HSSFWorkbook(inputStream);
			}
			//Reading sheet inside Workbook by its name
			sheetBook = excelWorkBook.getSheet(excelSheetName);
		} 	
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public String getDataExcel(String search) {
		System.out.println("::::: getDataExcel ("+search+") :::::");
		try {
			//Row 1
			row = sheetBook.getRow(0);
			int i = 0;
			Boolean stop = false;
			while (stop.equals(false)) {
				data = row.getCell(i).getStringCellValue();
				if (data.equals(search)) {
					//Takes next row
					row = sheetBook.getRow(1);
					try {
						cellData = row.getCell(i).getStringCellValue();
					} catch (java.lang.IllegalStateException e) {
						System.out.println("::::: Convertig Int to String :::::");
						cellNumberData = row.getCell(i).getNumericCellValue();
						cellData = Math.ceil(cellNumberData) == Math.floor(cellNumberData) ?  String.valueOf((int)row.getCell(i).getNumericCellValue()) : String.valueOf(row.getCell(i).getNumericCellValue());
					}
					row = sheetBook.getRow(0);	
					stop = true;
				}else {
					if((data.getBytes().length) < 0) {
						System.out.println("::::: Empty Column: "+data.getBytes().length+" :::::");
						stop = true;
					}
				}
				i++;	
			}
		} catch (Exception e) {System.out.println("¡¡¡¡¡ getDataExcel Method: "+e+"!!!!!");}
		return cellData;
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public String getDataExcel(String search, int scenario) {
		System.out.println("::::: getDataExcel ("+search+") scenario: "+scenario+" :::::");
		try {
			//Row 1
			row = sheetBook.getRow(0);
			int i = 0;
			Boolean stop = false;
			while (stop.equals(false)) {
				data = row.getCell(i).getStringCellValue();
				if (data.equals(search)) {
					//Takes next row
					row = sheetBook.getRow((1+scenario));
					try {
						cellData = row.getCell(i).getStringCellValue();
					} catch (java.lang.IllegalStateException e) {
						System.out.println("::::: Convertig Int to String :::::");
						cellNumberData = row.getCell(i).getNumericCellValue();
						cellData = Math.ceil(cellNumberData) == Math.floor(cellNumberData) ?  String.valueOf((int)row.getCell(i).getNumericCellValue()) : String.valueOf(row.getCell(i).getNumericCellValue());
					}
					System.out.println("::::: "+cellData);
					row = sheetBook.getRow(0);	
					stop = true;
				}else {
					if((data.getBytes().length) < 0) {
						System.out.println("::::: Empty Column: "+data.getBytes().length+" :::::");
						stop = true;
					}
				}
				i++;
			}
		} catch (Exception e) {System.out.println("¡¡¡¡¡ getDataExcel Method: "+e+"!!!!!");}
		return cellData;
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	@Override
	public String getDataExcel(String excelSheetName, String search, int scenario) {
		// TODO Auto-generated method stub
		return null;
	}
	@Override
	public void setDataExcel(String excelSheetName, String search, int scenario, int value) {
		// TODO Auto-generated method stub	
	}
}//Class