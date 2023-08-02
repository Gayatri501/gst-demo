package gstdemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class VendorRecordSendMail {
    public static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    public static boolean isCellBlank(Cell cell) {
        return cell.getCellType() == CellType.BLANK;
    }

    public static String getEmailForVendorCode(String filePath, String vendorCode) {
        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(file)) {
            Sheet sheet = workbook.getSheet("Sheet1");

            int vendorCodeColumnIndex = 0; 
            int emailColumnIndex = 1; 

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

    public static String getStringCellValue(Cell cell) {
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
    
    public static String getRowData(Row row) {
        StringBuilder rowData = new StringBuilder();
        int firstCellNum = row.getFirstCellNum();
        int lastCellNum = row.getLastCellNum();
        for (int cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                rowData.append(getStringCellValue(cell)).append("\t");
            }
        }
        return rowData.toString();
    }

    public static String getRowDataWithHeader(Row row) {
        StringBuilder rowData = new StringBuilder();
        int firstCellNum = row.getFirstCellNum();
        int lastCellNum = row.getLastCellNum();
        for (int cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                rowData.append(cell.getStringCellValue()).append("\t");
            }
        }
        return rowData.toString();
    }


    public static void sendEmail(String vendorCode, String records, String email) {
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
        String senderEmail = "gayatrikhairnar130@gmail.com"; // Replace with your sender email address
        String recipientEmail = email;

        // Sender's credentials
        String username = "gayatrikhairnar130@gmail.com"; // Replace with your sender email address
        String password = "tnrkdioeldbimdok"; // Replace with your sender email password

        // Create a session with authentication
        Session session = Session.getInstance(properties, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        });

        try {
           
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(senderEmail));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject("Mismatch Records for Vendor Code: " + vendorCode);

            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText("Please find the attached records for Vendor Code: " + vendorCode);

            Workbook attachmentWorkbook = new XSSFWorkbook();
            Sheet attachmentSheet = attachmentWorkbook.createSheet("Records");

            // Parse the records string and populate the cells in the attachment sheet
            String[] rows = records.split("\n");
            for (int i = 0; i < rows.length; i++) {
                String[] rowData = rows[i].split("\t");
                Row attachmentRow = attachmentSheet.createRow(i);
                for (int j = 0; j < rowData.length; j++) {
                    Cell attachmentCell = attachmentRow.createCell(j);
                    attachmentCell.setCellValue(rowData[j]);
                }
            }

            // Write the attachment workbook to a temporary file
            File tempFile = File.createTempFile("records", ".xlsx");
            tempFile.deleteOnExit();
            try (FileOutputStream fileOutputStream = new FileOutputStream(tempFile)) {
                attachmentWorkbook.write(fileOutputStream);
            }

          
            DataSource source = new FileDataSource(tempFile);

           
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(tempFile.getName());

            
            MimeMultipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);

           
            message.setContent(multipart);

          
            Transport.send(message);

            System.out.println("Email sent successfully to: " + email);
        } catch (MessagingException | IOException e) {
            e.printStackTrace();
        }
    }
    
}
