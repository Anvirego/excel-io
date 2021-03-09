package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 4.0 03/2021.
 * Excel: Reads data from an Excel File. 
 */
public final class Excel extends ExcelLogic {
	private static Excel excel = null;
	private String data;
	private Sheet sheetBook;
	private Row row;
	private String cellData;
	private double cellNumberData;
	
	public Excel(String excelFileName, String excelSheetName) throws FileNotFoundException, IOException {
		sheetBook = readDataExcel(excelSheetName, excelFileName);
	}//Constructor
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected static Excel getInstance(String excelFileName, String excelSheetName) throws FileNotFoundException, IOException {
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
}//Class