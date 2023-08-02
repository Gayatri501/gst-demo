package gstdemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Notin2B {
	public static void main(String[] args) {
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(
				"C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\GSTIN process input file.xlsx"))) {
			Sheet gstSheet = workbook.getSheet("GSTR2B");
			Sheet prSheet = workbook.getSheet("PURCHASE REGISTER");

			int gstColumnIndex = getColumnIndexByName(gstSheet, "GSTIN of supplier");
			int prColumnIndex = getColumnIndexByName(prSheet, "GSTIN of supplier");

			Set<String> gstUniqueIds = new HashSet<>();

			for (Row row : gstSheet) {
				Cell idCell = row.getCell(gstColumnIndex);
				if (idCell != null && idCell.getCellType() == CellType.STRING) {
					String id = idCell.getStringCellValue();
					gstUniqueIds.add(id);
				}
			}

			Sheet notgst2bSheet = workbook.getSheet("Not in GST2BR");
			int rowIndex = 0;

			if (notgst2bSheet == null) {
				notgst2bSheet = workbook.createSheet("Not in GST2BR");
			} else {
				workbook.removeSheetAt(workbook.getSheetIndex(notgst2bSheet));
				notgst2bSheet = workbook.createSheet("Not in GST2BR");
			}

			// Compare unique IDs from the GSTR2B sheet with the Purchase Register sheet
			for (Row row : prSheet) {
				Cell idCell = row.getCell(prColumnIndex);
				if (idCell != null && idCell.getCellType() == CellType.STRING) {
					String id = idCell.getStringCellValue();
					if (!gstUniqueIds.contains(id)) {
						// Unique ID not found in the Purchase Register sheet
						Row gstRow = gstSheet.getRow(row.getRowNum()); // Get the corresponding row from PR sheet
						if (gstRow != null) {
							copyRow(gstRow, notgst2bSheet.createRow(rowIndex++)); // Copy the entire row to "Not in PR"
																					// sheet
						}
					}
				}
			}

			// Save the workbook to a file
			try (FileOutputStream fileOut = new FileOutputStream(
					"C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\GSTIN process input file.xlsx")) {
				workbook.write(fileOut);
				System.out.println("Excel file updated with 'Not in GST2BR' sheet.");
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

	private static void copyRow(Row sourceRow, Row destinationRow) {
		for (Cell sourceCell : sourceRow) {
			int columnIndex = sourceCell.getColumnIndex();
			Cell destinationCell = destinationRow.createCell(columnIndex);

			// Copy cell value and style
			CellType cellType = sourceCell.getCellType();
			switch (cellType) {
			case BOOLEAN:
				destinationCell.setCellValue(sourceCell.getBooleanCellValue());
				break;
			case NUMERIC:
				destinationCell.setCellValue(sourceCell.getNumericCellValue());
				break;
			case STRING:
				destinationCell.setCellValue(sourceCell.getStringCellValue());
				break;
			case FORMULA:
				destinationCell.setCellFormula(sourceCell.getCellFormula());
				break;
			default:
				break;
			}

			// Copy cell style if available
			if (sourceCell.getCellStyle() != null) {
				CellStyle style = destinationRow.getSheet().getWorkbook().createCellStyle();
				style.cloneStyleFrom(sourceCell.getCellStyle());
				destinationCell.setCellStyle(style);
			}
		}
	}
}
