package gstdemo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class VendorCodeMatching {


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

    @SuppressWarnings("incomplete-switch")
    private static boolean compareCellValues(Cell cell1, Cell cell2) {
        if (cell1 != null && cell2 != null) {
            if (cell1.getCellType() == cell2.getCellType()) {
                switch (cell1.getCellType()) {
                    case STRING:
                        return cell1.getStringCellValue().equals(cell2.getStringCellValue());
                    case NUMERIC:
                        return cell1.getNumericCellValue() == cell2.getNumericCellValue();
                    case BOOLEAN:
                        return cell1.getBooleanCellValue() == cell2.getBooleanCellValue();
                    case FORMULA:
                        return cell1.getCellFormula().equals(cell2.getCellFormula());
                }
            }
        }
        return false;
    }

    public static void deleteRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    } 
    public static void findAndCopyMatchedRecords(Sheet sheet1, Sheet sheet2, int columnIndex1, int columnIndex2, Sheet matchedSheet) {
        int matchedRowIndex = 1; // Start from the first row in the matched sheet
        for (int rowIndex1 = 1; rowIndex1 <= sheet1.getLastRowNum(); rowIndex1++) {
            Row row1 = sheet1.getRow(rowIndex1);
            Cell vendorCodeCell1 = row1.getCell(columnIndex1);

            for (int rowIndex2 = 1; rowIndex2 <= sheet2.getLastRowNum(); rowIndex2++) {
                Row row2 = sheet2.getRow(rowIndex2);
                Cell vendorCodeCell2 = row2.getCell(columnIndex2);

                if (compareCellValues(vendorCodeCell1, vendorCodeCell2)) {
                    Row matchedRow = matchedSheet.createRow(matchedRowIndex);
                    copyRow(row1, matchedRow);
                    matchedRowIndex++;
                    break;
                }
            }
        }
    }
}
