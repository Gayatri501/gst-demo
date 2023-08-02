package gstdemo;

import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.File;

public class Mail {
	public static void main(String args[]) {
	
		String message = "Hello Dear, This is Message for security check";
		String subject = "WriterInformation : Security Check";
		String to = "shraddha.chaudhari@writerinformation.com";
		String from = "gayatrikhairnar130@gmail.com";

		// sendEmail(message,subject,to,from);
		sendAttachment(message, subject, to, from);
	}

	public static void sendAttachment(String message, String subject, String to, String from) {

		//String userDirectory = System.getProperty("user.dir");
		// variable for GMail Host
		String host = "smtp.gmail.com";

		// get the system properties
		Properties properties = System.getProperties();
		System.out.println("PROPERTIES " + properties);

		// setting important information to properties object

		// host set
		properties.put("mail.smtp.host", host);
		properties.put("mail.smtp.port", "587");
		properties.put("mail.smtp.ssl.enable", "true");
		properties.put("mail.smtp.auth", "true");

		// Step 1 : to get the session object

		Session session = Session.getInstance(properties, new Authenticator() {

			@Override
			protected PasswordAuthentication getPasswordAuthentication() {
				// TODO Auto-generated method stub
				return new PasswordAuthentication("gayatrikhairnar130@gmail.com", "tnrkdioeldbimdok");
			}

		});

		session.setDebug(true);

		// Step 2 : compose the message [text,multi media]
		MimeMessage m = new MimeMessage(session);

		// from email id
		try {

			// from email
			m.setFrom(from);

			// adding recipient to message
			m.addRecipient(javax.mail.Message.RecipientType.TO, new InternetAddress(to));

			// adding subject to message
			m.setSubject(subject);

			// attachement
			String path =  "C:\\Users\\Gayatri Khairnar\\Documents\\Java (selenium)\\selenium notes.txt";

			MimeMultipart mimeMultipart = new MimeMultipart();

			MimeBodyPart textMime = new MimeBodyPart();
			MimeBodyPart fileMime = new MimeBodyPart();

			try {
				textMime.setText(message);
				File file = new File(path);
				fileMime.attachFile(file);

				mimeMultipart.addBodyPart(textMime);
				mimeMultipart.addBodyPart(fileMime);

			} catch (Exception e) {
				e.printStackTrace();

			}

			m.setContent(mimeMultipart);
			// send
//
//			// Step 3 : Send Message using Transport class
//			Transport.send(m);
//			System.out.println("Message sent Success..........");

		} catch (MessagingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}


	}

	public static void sendEmail(String message, String subject, String to, String from) {

		// variable for GMail Host
		String host = "smtp.gmail.com";

		// get the system properties
		Properties properties = System.getProperties();
		System.out.println("PROPERTIES " + properties);

		// setting important information to properties object

		// host set
		properties.put("mail.smtp.host", host);
		properties.put("mail.smtp.port", "587 ");
		properties.put("mail.smtp.ssl.enable", "true");
		properties.put("mail.smtp.auth", "true");

		// Step 1 : to get the session object

		Session session = Session.getInstance(properties, new Authenticator() {

			@Override
			protected PasswordAuthentication getPasswordAuthentication() {
				// TODO Auto-generated method stub
				return new PasswordAuthentication("gayatrikhairnar130@gmail.com", "tnrkdioeldbimdok");
			}

		});

		session.setDebug(true);

		// Step 2 : compose the message [text,multi media]
		MimeMessage m = new MimeMessage(session);

		// from email id
		try {

			// from email
			m.setFrom(from);

			// adding recipient to message
			m.addRecipient(javax.mail.Message.RecipientType.TO, new InternetAddress(to));

			// adding subject to message
			m.setSubject(subject);

			// adding text to message
			m.setText(message);

			// send

			// Step 3 : Send Message using Transport class
			Transport.send(m);
			System.out.println("Message sent Success..........");

		} catch (MessagingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}
