package com.github.anvirego;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
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
 * ExcelIdem: Reads data from an Excel File. 
 */
public final class ExcelIdem implements ExcelInterface {
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
		readDataExcel();
	}//Constructor
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public static ExcelIdem getInstance(String excelFileName) throws FileNotFoundException, IOException {
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
	private void readDataExcel() throws FileNotFoundException, IOException {
		System.out.println("::::: readDataExcel :::::");
		File file = new File(excelFileName);
		//Create an object of FileInputStream to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		if(inputStream != null) {
			excelWorkBook = null;
			//Find the file extension by splitting excelFileName in substrings and getting only extension
			if((excelFileName.substring(excelFileName.indexOf("."))).equals(".xlsx")) {
				excelWorkBook = new XSSFWorkbook(inputStream);
			} else if((excelFileName.substring(excelFileName.indexOf("."))).equals(".xls")) {
				excelWorkBook = new HSSFWorkbook(inputStream);
			}
			
		} 	
	}//Method
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
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	@Override
	public String getDataExcel(String search) {
		// TODO Auto-generated method stub
		return null;
	}
	@Override
	public String getDataExcel(String search, int scenario) {
		// TODO Auto-generated method stub
		return null;
	}
}//Class