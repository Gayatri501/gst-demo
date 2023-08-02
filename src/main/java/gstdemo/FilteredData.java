package gstdemo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class FilteredData {
//    public static void main(String[] args) {
//        String filePath = "C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\Final output GSTIN file (13).xlsx";
//        String sheetName = "Final output sheet";
//        String columnName = "Mismatch Details";
//
//        try (FileInputStream file = new FileInputStream(filePath);
//             Workbook workbook = WorkbookFactory.create(file)) {
//            Sheet sheet = workbook.getSheet(sheetName);
//
//            // Find the column index of "Mismatch Details"
//            int columnIndex = findColumnIndex(sheet, columnName);
//
//            if (columnIndex == -1) {
//                System.out.println(columnName + " column not found in the sheet.");
//                return;
//            }
//
//            // Create a new workbook to store the filtered results
//            Workbook filteredWorkbook = WorkbookFactory.create(true);
//            Sheet filteredSheet = filteredWorkbook.createSheet("Filtered Results");
//
//            // Copy header row to the filtered sheet
//            Row headerRow = sheet.getRow(0);
//            Row filteredHeaderRow = filteredSheet.createRow(0);
//            copyRow(headerRow, filteredHeaderRow);
//
//            // Iterate over rows in the sheet to filter based on non-blank "Mismatch Details"
//            int filteredRowIndex = 1; // Start from the first row in the filtered sheet
//            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
//                Row row = sheet.getRow(rowIndex);
//                Cell cell = row.getCell(columnIndex);
//
//                // Check if the "Mismatch Details" cell is non-blank
//                if (cell != null && !isCellBlank(cell)) {
//                    // Copy the row to the filtered sheet
//                    Row filteredRow = filteredSheet.createRow(filteredRowIndex);
//                    copyRow(row, filteredRow);
//                    filteredRowIndex++;
//                }
//            }
//
//            // Save the filtered results to a new Excel file
//            try (FileOutputStream outFile = new FileOutputStream("C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\filtered_output.xlsx")) {
//                filteredWorkbook.write(outFile);
//            }
//
//            System.out.println("Non-blank records filtered successfully.");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    public static int findColumnIndex(Sheet sheet, String columnName) {
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

    public static void copyRow(Row sourceRow, Row destinationRow) {
        for (Cell sourceCell : sourceRow) {
            Cell destinationCell = destinationRow.createCell(sourceCell.getColumnIndex(), sourceCell.getCellType());
            copyCellValue(sourceCell, destinationCell);
        }
    }

    public static void copyCellValue(Cell sourceCell, Cell destinationCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                destinationCell.setCellValue(sourceCell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                destinationCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                destinationCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                destinationCell.setCellValue("");
        }
    }

    public static boolean isCellBlank(Cell cell) {
        if (cell.getCellType() == CellType.BLANK) {
            return true;
        } else if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {
            return true;
        }
        return false;
    }
    
    public static void deleteRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
}

