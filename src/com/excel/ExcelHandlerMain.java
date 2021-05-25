package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHandlerMain {
	private static int count=0;

	public void readExcel(String fileName, String sheetName) throws IOException {

		// Create an object of File class to open xlsx file

		File file = new File(fileName);

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workBook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			workBook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of HSSFWorkbook class

			workBook = new HSSFWorkbook(inputStream);

		}

		// Read sheet inside the workbook by its name

		Sheet sourceSheet = workBook.getSheet(sheetName);
		Sheet compareSheet = workBook.getSheet("compare");
		Sheet targetSheet = workBook.getSheet("target");

		// Find number of rows in excel file

		int rowCount = sourceSheet.getLastRowNum() - sourceSheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it

		for (int i = 0; i < rowCount + 1; i++) {

			Row row = sourceSheet.getRow(i);

			// Create a loop to print cell values in a row

			for (int j = 0; j < row.getLastCellNum(); j++) {

				// Print Excel data in console

				boolean flag = checkIfSourceFieldMatching(row.getCell(j).getStringCellValue(), compareSheet,targetSheet);

			//	System.out.print(row.getCell(j).getStringCellValue() + " : " + flag);

			}

			System.out.println();
		}
		inputStream.close();
		
		
		File targetFile = new File("target.xlsx");

		FileOutputStream outputStream = new FileOutputStream(targetFile);

		// write data in the excel file

		workBook.write(outputStream);

		// close output stream

		outputStream.close();

	}

	// Main function is calling readExcel function to read data from excel file

	private boolean checkIfSourceFieldMatching(String sourceCellValue, Sheet compareSheet,Sheet targetSheet) {

		int rowCount = compareSheet.getLastRowNum() - compareSheet.getFirstRowNum();
		for (int i = 0; i < rowCount + 1; i++) {

			Row compareRow = compareSheet.getRow(i);

			for (int j = 0; j < 1; j++) {
				
				if(sourceCellValue.equals(compareRow.getCell(j).getStringCellValue())) {
					Row newRow = targetSheet.createRow(count++);
					Cell payorColumnCell = newRow.createCell(0);
					Cell spsColumnCell = newRow.createCell(1);
					
					
					payorColumnCell.setCellValue(sourceCellValue);
					
					spsColumnCell.setCellValue(compareRow.getCell(j+1).getStringCellValue());
					spsColumnCell.setCellStyle(compareRow.getCell(j+1).getCellStyle());
					
				}

			}
		}

		return true;
	}

	public static void main(String... strings) throws IOException {

		// Create an object of ReadExcel class

		ExcelHandlerMain objExcelFile = new ExcelHandlerMain();

		// Prepare the path of excel file

		// Call read file method of the class to read data

		objExcelFile.readExcel("Source.xlsx", "Source_Sheet");

	}

}