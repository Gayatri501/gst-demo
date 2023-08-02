package gstdemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DeleteWriterColumn {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\Gayatri Khairnar\\Desktop\\Java Project\\gstdemo\\GST_Recon_Files\\Input\\matched_records.xlsx";
        String sheetName = "Matched Records";
        int columnIndex = 2; // Column index to delete (starting from 0)

        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {

            Sheet sheet = workbook.getSheet(sheetName);

            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    row.removeCell(cell);
                    shiftCellsLeft(row, columnIndex);
                }
            }

            try (FileOutputStream outFile = new FileOutputStream(filePath)) {
                workbook.write(outFile);
            }

            System.out.println("Column deleted successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void shiftCellsLeft(Row row, int columnIndex) {
        int lastCellIndex = row.getLastCellNum() - 1;
        for (int i = columnIndex; i < lastCellIndex; i++) {
            Cell currentCell = row.getCell(i);
            Cell nextCell = row.getCell(i + 1);
            if (currentCell != null) {
                if (nextCell != null) {
                    CellType nextCellType = nextCell.getCellType();
                    if (nextCellType == CellType.NUMERIC) {
                        currentCell.setCellValue(nextCell.getNumericCellValue());
                    } else if (nextCellType == CellType.STRING) {
                        currentCell.setCellValue(nextCell.getStringCellValue());
                    } else if (nextCellType == CellType.FORMULA) {
                        currentCell.setCellFormula(nextCell.getCellFormula());
                    } else if (nextCellType == CellType.BOOLEAN) {
                        currentCell.setCellValue(nextCell.getBooleanCellValue());
                    } else if (nextCellType == CellType.ERROR) {
                        currentCell.setCellErrorValue(nextCell.getErrorCellValue());
                    }
                } else {
                    currentCell.setCellValue("");
                }
            }
        }

        // Remove the last cell
        Cell lastCell = row.getCell(lastCellIndex);
        if (lastCell != null) {
            row.removeCell(lastCell);
        }
    }
}


