package gst.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import gstdemo.FilteredData;

public class FilteredDataTest {
	 public static void main(String[] args) {
		 
		 

		 String userDirectory = System.getProperty("user.dir");
			
			String filePath = userDirectory + "\\GST_Recon_Files\\Input\\Final output GSTIN file (13).xlsx";
	        String sheetName = "Final output sheet";
	        String columnName = "Mismatch Details";

	        
	        try (FileInputStream file = new FileInputStream(filePath);
	              Workbook workbook = WorkbookFactory.create(file)) {
	             Sheet sheet = workbook.getSheet(sheetName);
	 
	         	// Delete the row by index in sheet1
	 			int rowIndexToDelete = 1; // Specify the index of the row to delete
	 			FilteredData.deleteRow(sheet, rowIndexToDelete);

	             
	             // Find the column index of "Mismatch Details"
	             int columnIndex = FilteredData.findColumnIndex(sheet, columnName);
	 
	            if (columnIndex == -1) {
	                 System.out.println(columnName + " column not found in the sheet.");
	                 return;
	             }
	 
	             // Create a new workbook to store the filtered results
	             Workbook filteredWorkbook = WorkbookFactory.create(true);
	             Sheet filteredSheet = filteredWorkbook.createSheet("Filtered Results");
	 
	             // Copy header row to the filtered sheet
	             Row headerRow = sheet.getRow(0);
	             Row filteredHeaderRow = filteredSheet.createRow(0);
	             FilteredData.copyRow(headerRow, filteredHeaderRow);
	 
	             // Iterate over rows in the sheet to filter based on non-blank "Mismatch Details"
	             int filteredRowIndex = 1; // Start from the first row in the filtered sheet
	             for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
	                 Row row = sheet.getRow(rowIndex);
	                 Cell cell = row.getCell(columnIndex);
	 
	                 // Check if the "Mismatch Details" cell is non-blank
	                 if (cell != null && !FilteredData.isCellBlank(cell)) {
	                     // Copy the row to the filtered sheet
	                     Row filteredRow = filteredSheet.createRow(filteredRowIndex);
	                     FilteredData.copyRow(row, filteredRow);
	                     filteredRowIndex++;
	                 }
	             }
	 
             // Save the filtered results to a new Excel file
	             try (FileOutputStream outFile = new FileOutputStream(userDirectory + "\\GST_Recon_Files\\output\\filtered_output.xlsx")) {
	                 filteredWorkbook.write(outFile);
	             }
	 
	             System.out.println("Non-blank records filtered successfully.");
	         } catch (IOException e) {
	             e.printStackTrace();
	         }
	 }
}
