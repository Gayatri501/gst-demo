package usecases_related_chrome;

import java.io.BufferedInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.net.URLConnection;

public class DownloadFileUsingDriveLink {

	public static void main(String[] args) {
		String fileUrl = "https://drive.google.com/file/d/1PsbLBiyYhppArEYqSfYZSZffnBa6jmE9/view?usp=drive_link";
		String destinationFileName = "downloadedFile.txt";

		try {
			downloadFileFromDrive(fileUrl, destinationFileName);
			System.out.println("File downloaded successfully: " + destinationFileName);
		} catch (IOException e) {
			System.out.println("Error downloading the file: " + e.getMessage());
		}
	}

	private static void downloadFileFromDrive(String fileUrl, String destinationFileName) throws IOException {

		@SuppressWarnings("deprecation")
		URL url = new URL(fileUrl);
		URLConnection conn = url.openConnection();
		BufferedInputStream inputStream = new BufferedInputStream(conn.getInputStream());
		FileOutputStream outputStream = new FileOutputStream(destinationFileName);

		byte[] buffer = new byte[1024];
		int bytesRead;
		while ((bytesRead = inputStream.read(buffer)) != -1) {
			outputStream.write(buffer, 0, bytesRead);
		}

		outputStream.close();
		inputStream.close();
	}
}
