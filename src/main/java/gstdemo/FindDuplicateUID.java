package gstdemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FindDuplicateUID {
    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream("C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\uniqueid.xlsm"))) {
            Sheet sheet = workbook.getSheet("GSTIN");

            int uidColumnIndex = getColumnIndexByName(sheet, "UID");
            Set<String> uniqueIds = new HashSet<>();
            Set<String> duplicateIds = new HashSet<>();

            // Find duplicate IDs and store them in the duplicateIds set
            for (Row row : sheet) {
                Cell idCell = row.getCell(uidColumnIndex);
                if (idCell != null && idCell.getCellType() == CellType.STRING) {
                    String id = idCell.getStringCellValue();
                    if (!uniqueIds.add(id)) {
                        duplicateIds.add(id);
                    }
                }
            }

            // Create a new sheet to write duplicate IDs and their corresponding rows
            Sheet duplicateSheet = workbook.createSheet("duplicateUID");
            int rowIndex = 0;

            // Write the duplicate IDs and their corresponding rows to the new sheet
            for (Row row : sheet) {
                Cell idCell = row.getCell(uidColumnIndex);
                if (idCell != null && idCell.getCellType() == CellType.STRING) {
                    String id = idCell.getStringCellValue();
                    if (duplicateIds.contains(id)) {
                        Row newRow = duplicateSheet.createRow(rowIndex++);
                        for (int i = 0; i < row.getLastCellNum(); i++) {
                            Cell newCell = newRow.createCell(i);
                            Cell oldCell = row.getCell(i);
                            if (oldCell != null) {
                                CellType cellType = oldCell.getCellType();
                                switch (cellType) {
                                    case STRING:
                                        newCell.setCellValue(oldCell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        newCell.setCellValue(oldCell.getNumericCellValue());
                                        break;
                                    case BOOLEAN:
                                        newCell.setCellValue(oldCell.getBooleanCellValue());
                                        break;
                                    case FORMULA:
                                        newCell.setCellFormula(oldCell.getCellFormula());
                                        break;
                                    // Handle other cell types as needed
                                }
                            }
                        }
                    }
                }
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\uniqueid.xlsm")) {
                workbook.write(fileOut);
                System.out.println("Excel file updated with 'duplicateUID' sheet.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static int getColumnIndexByName(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getCellType() == CellType.STRING && columnName.equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column name not found: " + columnName);
    }
}
