package gstdemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.activation.DataHandler;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class VendorRecords {
    public static void main(String[] args) {

        String userDirectory = System.getProperty("user.dir");
        String filePath = userDirectory + "\\GST_Recon_Files\\output\\matched_records.xlsx";
        String sheetName = "Matched Records";
        String columnName = "Vendor Code";

        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet(sheetName);

            int columnIndex = findColumnIndex(sheet, columnName);

            if (columnIndex == -1) {
                System.out.println(columnName + " column not found in the sheet.");
                return;
            }

            Map<String, StringBuilder> vendorRecords = new HashMap<>();

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);

                if (cell != null && !isCellBlank(cell)) {
                    String vendorCode;
                    if (cell.getCellType() == CellType.NUMERIC) {
                        vendorCode = String.valueOf(cell.getNumericCellValue());
                    } else {
                        vendorCode = cell.getStringCellValue();
                    }

                    StringBuilder records = vendorRecords.get(vendorCode);
                    if (records == null) {
                        records = new StringBuilder();
                        vendorRecords.put(vendorCode, records);
                    }

                    records.append(getRowData(row)).append("\n");
                }
            }

            // Send emails for each vendor code
            for (Map.Entry<String, StringBuilder> entry : vendorRecords.entrySet()) {
                String vendorCode = entry.getKey();
                String records = entry.getValue().toString();

                // Get email address for vendor code
                String email = getEmailForVendorCode(userDirectory + "\\GST_Recon_Files\\Input\\EmailSheet.xlsx", vendorCode);

                // Send email with the combined records as attachment
                sendEmail(vendorCode, records, email);
            }

            try (FileOutputStream outFile = new FileOutputStream(userDirectory + "\\GST_Recon_Files\\output\\filtered_output1.xlsx")) {
                workbook.write(outFile);
            }

            System.out.println("Filtered records created and emails sent successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    private static boolean isCellBlank(Cell cell) {
        return cell.getCellType() == CellType.BLANK;
    }

    private static String getRowData(Row row) {
        StringBuilder rowData = new StringBuilder();
        for (Cell cell : row) {
            rowData.append(cell).append("\t");
        }
        return rowData.toString();
    }

    private static String getEmailForVendorCode(String filePath, String vendorCode) {
        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet("Sheet1");

            int vendorCodeColumnIndex = 0; // Specify the column index of the vendor code column
            int emailColumnIndex = 1; // Specify the column index of the email column

            for (Row row : sheet) {
                Cell vendorCodeCell = row.getCell(vendorCodeColumnIndex);
                Cell emailCell = row.getCell(emailColumnIndex);
                String code = getStringCellValue(vendorCodeCell);
                if (code.equalsIgnoreCase(vendorCode)) {
                    return getStringCellValue(emailCell);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellType();
        if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cellType == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else if (cellType == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cellType == CellType.FORMULA) {
            return cell.getCellFormula();
        } else if (cellType == CellType.BLANK) {
            return "";
        } else {
            return "";
        }
    }

    private static String getRowDataWithHeader(Row row, Row headerRow) {
        StringBuilder rowData = new StringBuilder();
        int firstCellNum = row.getFirstCellNum();
        int lastCellNum = row.getLastCellNum();
        for (int cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                Cell headerCell = headerRow.getCell(cellIndex);
                rowData.append(getStringCellValue(headerCell)).append(": ").append(cell).append("\t");
            }
        }
        return rowData.toString();
    }
    private static void sendEmail(String vendorCode, String records, String email) {
        System.out.println("Sending email to: " + email);
        System.out.println("Vendor Code: " + vendorCode);
        System.out.println("Email content:\n" + records);
        System.out.println();

        
        Properties properties = new Properties();
        properties.put("mail.smtp.host", "smtp.gmail.com");
        properties.put("mail.smtp.port", "587");
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.starttls.enable", "true");

        // Sender and recipient email addresses
        String senderEmail = "gayatrikhairnar130@gmail.com"; 
        String recipientEmail = email;

        // Sender's credentials
        String username = "gayatrikhairnar130@gmail.com";  
        String password = "tnrkdioeldbimdok";  

       
        Session session = Session.getInstance(properties, new javax.mail.Authenticator() {
            protected javax.mail.PasswordAuthentication getPasswordAuthentication() {
                return new javax.mail.PasswordAuthentication(username, password);
            }
        });

        try {
            // Create a new email message
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(senderEmail));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject("MisMatch Records for Vendor Code: " + vendorCode);

            // Create the email body with text and attachment
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText("Please find the attached records for Vendor Code: " + vendorCode);

            // Create a temporary file for the records
            File tempFile = File.createTempFile("records", ".xlsx");
            tempFile.deleteOnExit();

            // Write the records to the temporary file
            try (FileOutputStream fileOutputStream = new FileOutputStream(tempFile)) {
                fileOutputStream.write(records.getBytes());
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Create a DataSource from the temporary file
            DataSource source = new FileDataSource(tempFile);

            // Attach the DataSource to the email
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(tempFile.getName());

            // Create a multipart message to hold the body and attachment
            MimeMultipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);

            // Set the multipart as the email content
            message.setContent(multipart);

            // Send the email
            Transport.send(message);

            System.out.println("Email sent successfully to: " + email);
        } catch (MessagingException | IOException e) {
            e.printStackTrace();
        }
    }
}
