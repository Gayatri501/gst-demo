package gstdemo;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellUnMerge {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream("C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\formated1copy.xlsx");
            Workbook workbook = new XSSFWorkbook(file);

 
            String sheetName = "B2B";

            Sheet sheet = workbook.getSheet(sheetName);

            int rowIndex = 0; 

            // Shift the existing rows down to make room for the new row
            sheet.shiftRows(rowIndex, sheet.getLastRowNum(), 1);

           
            Row newRow = sheet.createRow(rowIndex);

          
            Cell cell1 = newRow.createCell(0);
            cell1.setCellValue("GSTIN of supplier");

            Cell cell2 = newRow.createCell(1);
            cell2.setCellValue("Trade/Legal name");

            Cell cell3 = newRow.createCell(2);
            cell3.setCellValue("Invoice number");

            Cell cell4 = newRow.createCell(3);
            cell4.setCellValue("Invoice type");

            Cell cell5 = newRow.createCell(4);
            cell5.setCellValue("Invoice Date");

            Cell cell6 = newRow.createCell(5);
            cell6.setCellValue("Invoice Value(₹)");

            Cell cell7 = newRow.createCell(6);
            cell7.setCellValue("Place of supply");

            Cell cell8 = newRow.createCell(7);
            cell8.setCellValue("Supply Attract Reverse Charge");

            
            Cell cell9 = newRow.createCell(8);
            cell9.setCellValue("Rate(%)");

            Cell cell10 = newRow.createCell(9);
            cell10.setCellValue("Taxable Value (₹)");

            Cell cell11 = newRow.createCell(10);
            cell11.setCellValue("Integrated Tax(₹)");

            Cell cell12 = newRow.createCell(11);
            cell12.setCellValue("Central Tax(₹)");

            Cell cell13 = newRow.createCell(12);
            cell13.setCellValue("State/UT Tax(₹)");

            Cell cell14 = newRow.createCell(13);
            cell14.setCellValue("Cess(₹)");
            
            Cell cell15 = newRow.createCell(14);
            cell15.setCellValue("GSTR-1/IFF/GSTR-5 Period");

            Cell cell16 = newRow.createCell(15);
            cell16.setCellValue("GSTR-1/IFF/GSTR-5 Filing Date");

            Cell cell17 = newRow.createCell(16);
            cell17.setCellValue("ITC Availability");

            Cell cell18 = newRow.createCell(17);
            cell18.setCellValue("Reason");

            Cell cell19 = newRow.createCell(18);
            cell19.setCellValue("Applicable % of Tax Rate");

            Cell cell20 = newRow.createCell(19);
            cell20.setCellValue("Source");

            Cell cell21 = newRow.createCell(20);
            cell21.setCellValue("IRN");

            Cell cell22 = newRow.createCell(21);
            cell22.setCellValue("IRN Date");

            
            FileOutputStream outputFile = new FileOutputStream("C:\\Users\\Gayatri Khairnar\\Desktop\\GST_Recon_Files\\B2Bformate.xlsx");
            workbook.write(outputFile);
            outputFile.close();

            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
