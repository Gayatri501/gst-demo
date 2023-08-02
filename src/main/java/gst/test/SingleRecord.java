
package gst.test;
 
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import gstdemo.FilteredData;

public class SingleRecord {
    public static void main(String[] args) {

        String userDirectory = System.getProperty("user.dir");
        String filePath = userDirectory + "\\GST_Recon_Files\\output\\matched_records.xlsx";
        String sheetName = "Matched Records";
        String columnName = "Vendor Code";

        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet(sheetName);

           
            int columnIndex = FilteredData.findColumnIndex(sheet, columnName);

            if (columnIndex == -1) {
                System.out.println(columnName + " column not found in the sheet.");
                return;
            }

            Map<String, Sheet> filteredSheets = new HashMap<>();

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);

                if (cell != null && !FilteredData.isCellBlank(cell)) {
                    String vendorCode;
                    if (cell.getCellType() == CellType.NUMERIC) {
                        
                        vendorCode = String.valueOf(cell.getNumericCellValue());
                    } else {
                        
                        vendorCode = cell.getStringCellValue();
                    }

                    Sheet filteredSheet = filteredSheets.get(vendorCode);
                    if (filteredSheet == null) {
                       
                        filteredSheet = workbook.createSheet("Vendor Code " + vendorCode);
                        filteredSheets.put(vendorCode, filteredSheet);

                     
                        Row headerRow = sheet.getRow(0);
                        Row filteredHeaderRow = filteredSheet.createRow(0);
                        FilteredData.copyRow(headerRow, filteredHeaderRow);
                    }

                    // Copy the row to the filtered sheet
                    Row filteredRow = filteredSheet.createRow(filteredSheet.getLastRowNum() + 1);
                    FilteredData.copyRow(row, filteredRow);
                }
            }

            // Process each filtered sheet
            for (Sheet filteredSheet : filteredSheets.values()) {
              
                //sendEmail(filteredSheet);
            }

            try (FileOutputStream outFile = new FileOutputStream(userDirectory + "\\GST_Recon_Files\\output\\filtered_output1.xlsx")) {
                workbook.write(outFile);
            }

            System.out.println("Filtered records created and emails sent successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void sendEmail(Sheet sheet) {
     
        System.out.println("Sending email for Vendor Code: " + sheet.getSheetName());
        System.out.println("Email content:");
        for (Row row : sheet) {
            for (Cell cell : row) {
                // Process and retrieve cell data as needed
                String cellValue = cell.toString();
                System.out.print(cellValue + " ");
            }
            System.out.println();
        }
        System.out.println();
    }
}
