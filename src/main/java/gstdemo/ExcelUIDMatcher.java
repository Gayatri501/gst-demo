package gstdemo;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUIDMatcher {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\UID Matching Input.xlsx";
        String sheet1Name = "PR";
        String sheet2Name = "GSTIN";

        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet1 = workbook.getSheet(sheet1Name);
            Sheet sheet2 = workbook.getSheet(sheet2Name);

            // Find the UID column index in both sheets
            int uidColumnIndex1 = findColumnIndex(sheet1, "UID");
            int uidColumnIndex2 = findColumnIndex(sheet2, "UID");

            if (uidColumnIndex1 == -1 || uidColumnIndex2 == -1) {
                System.out.println("UID column not found in one or both sheets.");
                return;
            }

            // Add or update the 'Status' column header in the first row
            addOrUpdateColumnHeader(sheet1, "Status", uidColumnIndex1 + 1);
            addOrUpdateColumnHeader(sheet2, "Status", uidColumnIndex2 + 1);

            // Iterate over rows in the 'Purchase register' sheet
            for (Row row1 : sheet1) {
                Cell uidCell1 = row1.getCell(uidColumnIndex1);
                if (uidCell1 != null) {
                    String uidValue = uidCell1.getStringCellValue();

                    // Iterate over rows in the 'GSTR2B' sheet to find a match
                    boolean matchFound = false;
                    for (Row row2 : sheet2) {
                        Cell uidCell2 = row2.getCell(uidColumnIndex2);
                        if (uidCell2 != null && uidCell2.getStringCellValue().equals(uidValue)) {
                            matchFound = true;
                            break;
                        }
                    }

                    // Update the 'Status' column based on the match result
                    Cell statusCell1 = addOrUpdateCell(row1, uidColumnIndex1 + 1);
                    statusCell1.setCellValue(matchFound ? "Match" : "Mismatch");
                }
            }

            // Iterate over rows in the 'GSTR2B' sheet to find mismatches
            for (Row row2 : sheet2) {
                Cell uidCell2 = row2.getCell(uidColumnIndex2);
                if (uidCell2 != null) {
                    String uidValue = uidCell2.getStringCellValue();

                    // Iterate over rows in the 'Purchase register' sheet to find a match
                    boolean matchFound = false;
                    for (Row row1 : sheet1) {
                        Cell uidCell1 = row1.getCell(uidColumnIndex1);
                        if (uidCell1 != null && uidCell1.getStringCellValue().equals(uidValue)) {
                            matchFound = true;
                            break;
                        }
                    }

                    // Update the 'Status' column based on the match result
                    Cell statusCell2 = addOrUpdateCell(row2, uidColumnIndex2 + 1);
                    statusCell2.setCellValue(matchFound ? "Match" : "Mismatch");
                }
            }

            // Save the changes to the Excel file
            try (FileOutputStream outFile = new FileOutputStream(filePath)) {
                workbook.write(outFile);
            }

            System.out.println("UID matching completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Column not found
    }

    private static Cell addOrUpdateCell(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    private static void addOrUpdateColumnHeader(Sheet sheet, String columnName, int columnIndex) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            headerRow = sheet.createRow(0);
        }

        Cell headerCell = headerRow.getCell(columnIndex);
        if (headerCell == null) {
            headerCell = headerRow.createCell(columnIndex);
        }
        headerCell.setCellValue(columnName);
    }
}
