package com.regav;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.regav.interfaces.ExcelInterface;
/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 3.0 03/2021.
 * ExcelIdem: Excel tests. 
 */
public class ExcelTests {
	
	public static void main (String args[]) throws FileNotFoundException, IOException {
		ExcelInterface ei = ExcelIdem.getInstance("DataTable.xlsx");
		System.out.print("Data on Sheet 'MainInfo', \nColumn 'ENVIRONMENT' and \nROW '1': \n ::::: "+ei.getDataExcel("MainInfo", "ENVIRONMENT", 0)+" ::::: \n\n");
		
		System.out.print("Data on Sheet 'MainInfo', \nColumn 'ENVIRONMENT' and \nROW '3': \n ::::: "+ei.getDataExcel("MainInfo", "ENVIRONMENT", 2)+" ::::: \n\n");

		System.out.println("\n");
		
		ExcelInterface ei2 = Excel.getInstance("DataTable.xlsx", "Description");
		System.out.print("Data on Sheet 'Description', \nColumn 'REFERENCE' and \nROW '1': \n ::::: "+ei2.getDataExcel("REFERENCE")+" :::::\n\n");
	}//Main
}//Class