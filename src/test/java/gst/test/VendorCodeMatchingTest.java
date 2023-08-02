package gst.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import gstdemo.VendorCodeMatching;

public class VendorCodeMatchingTest {
	public static void main(String[] args) {

		String userDirectory = System.getProperty("user.dir");

		String filePath1 = userDirectory + "\\GST_Recon_Files\\output\\filtered_output.xlsx";
		String filePath2 = userDirectory + "\\GST_Recon_Files\\Input\\Master Sheet.xls";

		String sheetName1 = "Filtered Results";
		String sheetName2 = "Sheet1";
		String columnName = "Vendor Code";

		try (FileInputStream file1 = new FileInputStream(filePath1);
				FileInputStream file2 = new FileInputStream(filePath2);
				Workbook workbook1 = WorkbookFactory.create(file1);
				Workbook workbook2 = WorkbookFactory.create(file2)) {

			// Get the sheets from both workbooks
			Sheet sheet1 = workbook1.getSheet(sheetName1);
			Sheet sheet2 = workbook2.getSheet(sheetName2);

//			// Delete the row by index in sheet1
//			int rowIndexToDelete = 1; // Specify the index of the row to delete
//			VendorCodeMatching.deleteRow(sheet1, rowIndexToDelete);

			
			int columnIndex1 = VendorCodeMatching.findColumnIndex(sheet1, columnName);
			int columnIndex2 = VendorCodeMatching.findColumnIndex(sheet2, columnName);

			if (columnIndex1 == -1 || columnIndex2 == -1) {
				System.out.println(columnName + "column not found in one or both sheets.");
				return;
			}

			// Create a new workbook to store the matched records
			Workbook matchedWorkbook = WorkbookFactory.create(true);
			Sheet matchedSheet = matchedWorkbook.createSheet("Matched Records");

			// Copy header row to the matched sheet
			Row headerRow1 = sheet1.getRow(0);
			@SuppressWarnings("unused")
			Row headerRow2 = sheet2.getRow(0);
			Row matchedHeaderRow = matchedSheet.createRow(0);
			VendorCodeMatching.copyRow(headerRow1, matchedHeaderRow);

		
			VendorCodeMatching.findAndCopyMatchedRecords(sheet1, sheet2, columnIndex1, columnIndex2, matchedSheet);

			// Save the matched records to a new Excel file
			try (FileOutputStream outFile = new FileOutputStream(
					userDirectory + "\\GST_Recon_Files\\output\\matched_records.xlsx")) {
				matchedWorkbook.write(outFile);
			}

			System.out.println("Matched records copied successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
