package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 4.0 03/2021.
 * ExcelIdem: Reads data from an Excel File. 
 */
public final class ExcelIdem extends ExcelLogic {
	private static ExcelIdem excelI= null;
	private String excelFileName;
	private String data;
	private Sheet sheetBook;
	private static Workbook excelWorkBook;
	private Row row;
	private String cellData;
	private double cellNumberData;
	
	public ExcelIdem(String excelFileName) throws FileNotFoundException, IOException {
		this.excelFileName = excelFileName;
		excelWorkBook = readDataExcel(excelFileName);
	}//Constructor
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected static ExcelIdem getInstance(String excelFileName) throws FileNotFoundException, IOException {
		System.out.println("==== Get ExcelIdem Instance =====");
		if(excelI == null) {
			System.out.println("New Instance");
			excelI = new ExcelIdem(excelFileName);
			return excelI;
		} else {
			System.out.println("Old Instance");
			return excelI;
		}
	}
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public String getDataExcel(String excelSheetName, String search, int scenario) {
		System.out.println("::::: getDataExcel ("+search+") scenario: "+scenario+" :::::");
		try {
			//Reading sheet inside Workbook by its name
			sheetBook = excelWorkBook.getSheet(excelSheetName);
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
	public void setDataExcel(String excelSheetName, String search, int scenario, int value) {
		System.out.println("::::: setDataExcel ("+search+") scenario: "+scenario+" :::::");
		try {
			//Reading sheet inside Workbook by its name
			sheetBook = excelWorkBook.getSheet(excelSheetName);
			//Row 1
			row = sheetBook.getRow(0);
			int i = 0;
			Boolean stop = false;
			while (stop.equals(false)) {
				data = row.getCell(i).getStringCellValue();
				if (data.equals(search)) {
					//Takes next row
					row = sheetBook.getRow((1+scenario));
					row.getCell(i).setCellValue(value);
					System.out.println("::::: "+value);
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
			FileOutputStream fileOut = new FileOutputStream(excelFileName);
			excelWorkBook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {System.out.println("¡¡¡¡¡ setDataExcel Method: "+e+"!!!!!");}
	}//Method
	
}//Class