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

public class B2BFormate {
    public static void main(String[] args) {
        String sourceFilePath = "C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\042023_33AABCW6386P1ZM_GSTR2B_16052023.xlsx";
        String destinationFilePath = "C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\B2Bformate.xlsx";

        try (Workbook sourceWorkbook = new XSSFWorkbook(new FileInputStream(sourceFilePath));
             Workbook destinationWorkbook = new XSSFWorkbook()) {

            // Get the source sheet
            Sheet sourceSheet = sourceWorkbook.getSheet("B2B");
            if (sourceSheet == null) {
                System.out.println("Source sheet 'B2B' not found.");
                return;
            }

            // Create the destination sheet
            Sheet destinationSheet = destinationWorkbook.createSheet("B2B");

            // Copy the remaining rows from source to destination
            int rowIndex = 0;
            for (Row sourceRow : sourceSheet) {
                if (sourceRow.getRowNum() < 6) {
                    // Skip the first 6 rows
                    continue;
                }

                Row destinationRow = destinationSheet.createRow(rowIndex++);
                copyCells(sourceRow, destinationRow);
            }

            // Add cell values to the destination sheet
            Row headerRow = destinationSheet.createRow(0);
            Cell cell1 = headerRow.createCell(0);
            cell1.setCellValue("GSTIN of supplier");

            Cell cell2 = headerRow.createCell(1);
            cell2.setCellValue("Trade/Legal name");

            Cell cell3 = headerRow.createCell(2);
            cell3.setCellValue("Invoice number");

            Cell cell4 = headerRow.createCell(3);
            cell4.setCellValue("Invoice type");
            
            Cell cell5 = headerRow.createCell(4);
            cell5.setCellValue("Invoice Date");

            Cell cell6 = headerRow.createCell(5);
            cell6.setCellValue("Invoice Value(₹)");
            
            Cell cell7 = headerRow.createCell(6);
            cell7.setCellValue("Place of supply");

            Cell cell8 = headerRow.createCell(7);
            cell8.setCellValue("Supply Attract Reverse Charge");
            
            Cell cell9 = headerRow.createCell(8);
            cell9.setCellValue("Rate(%)");

            Cell cell10 = headerRow.createCell(9);
            cell10.setCellValue("Taxable Value (₹)");
            
            Cell cell11 = headerRow.createCell(10);
            cell11.setCellValue("Integrated Tax(₹)");

            Cell cell12 = headerRow.createCell(11);
            cell12.setCellValue("Central Tax(₹)");
            
            Cell cell13 = headerRow.createCell(12);
            cell13.setCellValue("State/UT Tax(₹)");

            Cell cell14 = headerRow.createCell(13);
            cell14.setCellValue("Cess(₹)");
            
            Cell cell15 = headerRow.createCell(14);
            cell15.setCellValue("GSTR-1/IFF/GSTR-5 Period");

            Cell cell16 = headerRow.createCell(15);
            cell16.setCellValue("GSTR-1/IFF/GSTR-5 Filing Date");

            Cell cell17 = headerRow.createCell(16);
            cell17.setCellValue("ITC Availability");

            Cell cell18 = headerRow.createCell(17);
            cell18.setCellValue("Reason");

            Cell cell19 = headerRow.createCell(18);
            cell19.setCellValue("Applicable % of Tax Rate");

            Cell cell20 = headerRow.createCell(19);
            cell20.setCellValue("Source");

            Cell cell21 = headerRow.createCell(20);
            cell21.setCellValue("IRN");

            Cell cell22 = headerRow.createCell(21);
            cell22.setCellValue("IRN Date");
            

            // Save the destination workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(destinationFilePath)) {
                destinationWorkbook.write(fileOut);
                System.out.println("Destination file created successfully.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void copyCells(Row sourceRow, Row destinationRow) {
        for (Cell sourceCell : sourceRow) {
            int columnIndex = sourceCell.getColumnIndex();
            Cell destinationCell = destinationRow.createCell(columnIndex);
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

            if (sourceCell.getCellComment() != null) {
                destinationCell.setCellComment(sourceCell.getCellComment());
            }

            if (sourceCell.getCellStyle() != null) {
                CellStyle style = destinationRow.getSheet().getWorkbook().createCellStyle();
                style.cloneStyleFrom(sourceCell.getCellStyle());
                destinationCell.setCellStyle(style);
            }
        }
    }
}
