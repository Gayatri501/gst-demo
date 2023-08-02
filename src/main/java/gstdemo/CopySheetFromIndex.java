package gstdemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopySheetFromIndex {
	public static void main(String[] args) {
		try {

			FileInputStream inputFile = new FileInputStream("C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\042023_33AABCW6386P1ZM_GSTR2B_16052023.xlsx");
			Workbook inputWorkbook = new XSSFWorkbook(inputFile);

			Workbook outputWorkbook = new XSSFWorkbook();

			String sheetName = "B2B";
			int startrow = 6;
			// Get the input sheet
			Sheet inputSheet = inputWorkbook.getSheet(sheetName);

			if (inputSheet != null) {
				// Create a new sheet in the output workbook with the same name
				Sheet outputSheet = outputWorkbook.createSheet(sheetName);

				copyRows(inputSheet, outputSheet, startrow);

				FileOutputStream outputFile = new FileOutputStream(
						"C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\formated.xlsx");
				outputWorkbook.write(outputFile);

				outputFile.close();
			}

			outputWorkbook.close();
			inputWorkbook.close();
			inputFile.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void copyRows(Sheet sourceSheet, Sheet targetSheet, int startRow) {
		int rowCount = sourceSheet.getLastRowNum();
		int rowIndex = 0;
		for (int i = startRow; i <= rowCount; i++) {
			Row sourceRow = sourceSheet.getRow(i);
			if (sourceRow != null) {
				Row targetRow = targetSheet.createRow(rowIndex++);
				copyCells(sourceRow, targetRow);
			}
		}
	}

	private static void copyCells(Row sourceRow, Row targetRow) {
		int columnCount = sourceRow.getLastCellNum();
		for (int j = 0; j < columnCount; j++) {
			Cell sourceCell = sourceRow.getCell(j);
			Cell targetCell = targetRow.createCell(j);
			copyCellValue(sourceCell, targetCell);
			copyCellStyle(sourceCell, targetCell, targetRow.getSheet().getWorkbook());
		}
	}

	@SuppressWarnings("deprecation")
	private static void copyCellValue(Cell sourceCell, Cell targetCell) {
		if (sourceCell == null) {
			targetCell.setCellType(CellType.BLANK);
		} else {
			CellType cellType = sourceCell.getCellType();

			if (cellType == CellType.STRING) {
				targetCell.setCellValue(sourceCell.getStringCellValue());
			} else if (cellType == CellType.NUMERIC) {
				targetCell.setCellValue(sourceCell.getNumericCellValue());
			} else if (cellType == CellType.BOOLEAN) {
				targetCell.setCellValue(sourceCell.getBooleanCellValue());
			} else if (cellType == CellType.FORMULA) {
				targetCell.setCellFormula(sourceCell.getCellFormula());
			}
		}
	}

	private static void copyCellStyle(Cell sourceCell, Cell targetCell, Workbook targetWorkbook) {
		if (sourceCell != null) {
			CellStyle sourceCellStyle = sourceCell.getCellStyle();
			CellStyle targetCellStyle = targetWorkbook.createCellStyle();
			targetCellStyle.cloneStyleFrom(sourceCellStyle);
			targetCell.setCellStyle(targetCellStyle);
		}
	}
	
	
	
}
