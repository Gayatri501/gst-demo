package gst.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DeleteColumn2BTest {
	 public static void main(String[] args) {
	        String sourceFilePath = "C:\\Users\\Gayatri Khairnar\\Desktop\\Java Project\\gstdemo\\GST_Recon_Files\\output\\matched_records.xlsx";
	        String sourceSheetName = "Matched Records";

	        String destinationFilePath = sourceFilePath.replace(".xlsx", "_AfterColumnDeleted.xlsx");
	        String destinationSheetName = "Matched Records";

	        int[] columnIndicesToCopy = {0, 1, 3, 4, 6, 7, 9, 10, 12, 13, 15, 16, 17, 20, 23, 26, 28, 30, 31, 32, 34, 35, 36, 38, 39, 40, 41, 42};

	        try (FileInputStream sourceFile = new FileInputStream(sourceFilePath);
	             Workbook sourceWorkbook = WorkbookFactory.create(sourceFile);
	             Workbook destinationWorkbook = new XSSFWorkbook()) { // Use XSSFWorkbook to create a new destination workbook

	            Sheet sourceSheet = sourceWorkbook.getSheet(sourceSheetName);
	            Sheet destinationSheet = destinationWorkbook.createSheet(destinationSheetName);

	            int rowIndex = 0;
	            for (Row sourceRow : sourceSheet) {
	                Row destinationRow = destinationSheet.createRow(rowIndex);

	                for (int i = 0; i < columnIndicesToCopy.length; i++) {
	                    int sourceColumnIndex = columnIndicesToCopy[i];
	                    int destinationColumnIndex = i;

	                    Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
	                    Cell destinationCell = destinationRow.createCell(destinationColumnIndex);

	                    if (sourceCell != null) {
	                        CellType cellType = sourceCell.getCellType();

	                        switch (cellType) {
	                            case STRING:
	                                destinationCell.setCellValue(sourceCell.getStringCellValue());
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
	                                break;
	                        }

	                        
	                        CellStyle sourceCellStyle = sourceCell.getCellStyle();
	                        CellStyle destinationCellStyle = destinationWorkbook.createCellStyle();
	                        destinationCellStyle.cloneStyleFrom(sourceCellStyle);
	                        destinationCell.setCellStyle(destinationCellStyle);
	                    }
	                }
	                rowIndex++;
	            }

	            try (FileOutputStream outputFile = new FileOutputStream(destinationFilePath)) {
	                destinationWorkbook.write(outputFile);
	            }

	            System.out.println("Sheet copied successfully!");
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
}
