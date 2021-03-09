package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.anvirego.interfaces.ExcelInterface;


/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 2.0 03/2021. ExcelTests: Examples of implementation.
 */
public class ExcelTests {

	public static void main(String args[]) throws FileNotFoundException, IOException {
		ExcelInterface ei3 = ExcelInstance.getInstance("DataTable.xlsx");

		System.out.print("Data on Sheet 'MainInfo', \nColumn 'ENVIRONMENT' and \nROW '1': \n ::::: "
				+ ei3.getDataExcel("MainInfo", "ENVIRONMENT", 0) + " ::::: \n\n");

		System.out.print("Data on Sheet 'MainInfo', \nColumn 'ENVIRONMENT' and \nROW '3': \n ::::: "
				+ ei3.getDataExcel("MainInfo", "ENVIRONMENT", 2) + " ::::: \n\n");

		System.out.println("\n");

		ei3 = ExcelInstance.getInstance("DataTable.xlsx", "Description");

		System.out.print("Data on Sheet 'Description', \nColumn 'REFERENCE' and \nROW '1': \n ::::: "
				+ ei3.getDataExcel("REFERENCE") + " :::::\n\n");

		System.out.print("Data on Sheet 'Description', \nColumn 'REFERENCE' and \nROW '4': \n ::::: "
				+ ei3.getDataExcel("REFERENCE", 3) + " :::::\n\n");

		System.out.println("\n");
	}// Main

}// Class