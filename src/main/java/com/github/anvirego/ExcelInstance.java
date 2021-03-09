package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.anvirego.interfaces.ExcelInterface;
/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 4.0 03/2021.
 * ExcelInstance: Defines and creates excel's instances. 
 */
public class ExcelInstance {

	public static ExcelInterface getExcelInstance(String excelFileName, String excelSheetName) throws FileNotFoundException, IOException {
		ExcelInterface ei = Excel.getInstance(excelFileName, excelSheetName);
		return ei;
	}
	
	public static ExcelInterface getExcelInstance(String excelFileName) throws FileNotFoundException, IOException {
		ExcelInterface ei = ExcelIdem.getInstance(excelFileName);
		return ei;
	}
}//Class
