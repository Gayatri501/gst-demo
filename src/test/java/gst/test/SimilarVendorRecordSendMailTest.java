package gst.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import gstdemo.VendorRecordSendMail;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SimilarVendorRecordSendMailTest {
	public static void main(String[] args) {

        String userDirectory = System.getProperty("user.dir");
        String filePath = userDirectory + "\\GST_Recon_Files\\output\\matched_records_AfterColumnDeleted.xlsx";
        String sheetName = "Matched Records";
        String columnName = "Vendor Code";

        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet(sheetName);

            int columnIndex = VendorRecordSendMail.findColumnIndex(sheet, columnName);

            if (columnIndex == -1) {
                System.out.println(columnName + " column not found in the sheet.");
                return;
            }

            Map<String, StringBuilder> vendorRecords = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            StringBuilder header = new StringBuilder(VendorRecordSendMail.getRowDataWithHeader(headerRow)).append("\n");
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);

                if (cell != null && !VendorRecordSendMail.isCellBlank(cell)) {
                    String vendorCode;
                    if (cell.getCellType() == CellType.NUMERIC) {
                        vendorCode = String.valueOf(cell.getNumericCellValue());
                    } else {
                        vendorCode = cell.getStringCellValue();
                    }

                    StringBuilder records = vendorRecords.get(vendorCode);
                    if (records == null) {
                        records = new StringBuilder(header.toString());
                        vendorRecords.put(vendorCode, records);
                    }

                    records.append(VendorRecordSendMail.getRowData(row)).append("\n");
                }
            }


            // Send emails for each vendor code
            for (Map.Entry<String, StringBuilder> entry : vendorRecords.entrySet()) {
                String vendorCode = entry.getKey();
                String records = entry.getValue().toString();

                // Get email address for vendor code
                String email = VendorRecordSendMail.getEmailForVendorCode(userDirectory + "\\GST_Recon_Files\\Input\\EmailSheet.xlsx", vendorCode);

              
                VendorRecordSendMail.sendEmail(vendorCode, records, email);
            }

            try (FileOutputStream outFile = new FileOutputStream(userDirectory + "\\GST_Recon_Files\\output\\filtered_output1.xlsx")) {
                workbook.write(outFile);
            }

            System.out.println("Filtered records created and emails sent successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
