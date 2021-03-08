package com.regav.interfaces;
/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 3.0 03/2021.
 * ExcelInterface: Interface Implementation. 
 */
public interface ExcelInterface {
	//Gets data by Column's Name.
	public String getDataExcel(String search);
		
	//Gets data by Columns's Name & Iterates over Column according to scenario's value.
	public String getDataExcel(String search, int scenario);
		
	//Gets data by Columns's Name & Iterates over Column according to scenario's value.
	public String getDataExcel(String excelSheetName, String search, int scenario);
		
	//Sets value by Columns's Name & Iterates over Column according to scenario's value.
	public void setDataExcel(String excelSheetName, String search, int scenario, int value);
}//Interface