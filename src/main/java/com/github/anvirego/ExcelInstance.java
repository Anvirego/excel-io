package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.anvirego.interfaces.ExcelInterface;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 1.0 03/2021. 
 * ExcelInstance: Defines and creates excel's instances to work with the Interface implemented.
 */
public class ExcelInstance {
	public static ExcelInterface getInstance(String excelFileName, String excelSheetName)
			throws FileNotFoundException, IOException {
		ExcelInterface ei = Excel.getInstance(excelFileName, excelSheetName);
		return ei;
	}// Method

	public static ExcelInterface getInstance(String excelFileName) throws FileNotFoundException, IOException {
		ExcelInterface ei = ExcelIdem.getInstance(excelFileName);
		return ei;
	}// Method

}// Class