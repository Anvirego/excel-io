package com.github.anvirego;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.anvirego.interfaces.ExcelInterface;
/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 2.0 03/2021.
 * ExcelLogic: Library main logic. 
 */
public class ExcelLogic implements ExcelInterface {
	private Sheet sheetBook;
	private static Workbook excelWorkBook;
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected Sheet readDataExcel(String excelSheetName, String excelFileName) throws FileNotFoundException, IOException {		
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
		return sheetBook;
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected Workbook readDataExcel(String excelFileName) throws FileNotFoundException, IOException {		
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
		return excelWorkBook;
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