package usecases_related_chrome;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;

import java.util.Properties;

public class DownloadFilePersonal {
	private static final String APPLICATION_NAME = "Writer Downloads";
	private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
	private static String TOKENS_DIRECTORY_PATH;
	private static String FOLDER_ID;
	private static String CLIENT_SECRET_PATH;
	private static String DESTINATION_FOLDER_PATH;

	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				createAndShowGUI();
			}
		});
	}

	private static void createAndShowGUI() {
		JFrame frame = new JFrame("Download Files from Drive");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setLayout(new FlowLayout());

		JButton downloadButton = new JButton("Download File");
		downloadButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					String userHome = System.getProperty("user.home");
					String configFilePath = userHome + File.separator + "config.properties";
					ConfigReader configReader = new ConfigReader(configFilePath);

					Properties properties = configReader.getProperties();
					TOKENS_DIRECTORY_PATH = properties.getProperty("tokens_directory_path");
					FOLDER_ID = properties.getProperty("folder_id");
					CLIENT_SECRET_PATH = properties.getProperty("client_secrets_path");
					DESTINATION_FOLDER_PATH = properties.getProperty("destination_folder_path");

					downloadFiles();
					JOptionPane.showMessageDialog(frame, "Files downloaded successfully!");
				} catch (IOException | GeneralSecurityException ex) {
					JOptionPane.showMessageDialog(frame, "Error occurred during download: " + ex.getMessage(), "Error",
							JOptionPane.ERROR_MESSAGE);
				}
			}
		});

		frame.add(downloadButton);
		frame.pack();
		frame.setVisible(true);
	}

//    private static Properties loadConfigProperties() throws IOException {
//        Properties properties = new Properties();
//        try (InputStream input = DownloadFilePersonal.class.getResourceAsStream("config.properties")) {
//            properties.load(input);
//        }
//        return properties;
//    }

	private static void downloadFiles() throws IOException, GeneralSecurityException {
		// Build the HTTP transport and set up the credentials
		HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
		Credential credential = authorize(httpTransport);

		// Create the output folder if it doesn't exist
		File outputFolder = new File(DESTINATION_FOLDER_PATH);
		if (!outputFolder.exists()) {
			if (outputFolder.mkdirs()) {
				System.out.println("Output folder created: " + DESTINATION_FOLDER_PATH);
			} else {
				throw new IOException("Failed to create the output folder: " + DESTINATION_FOLDER_PATH);
			}
		}

		downloadFilesFromFolder(httpTransport, credential, FOLDER_ID, DESTINATION_FOLDER_PATH);
	}

	private static Credential authorize(HttpTransport httpTransport) throws IOException {
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY,
				new InputStreamReader(new FileInputStream(CLIENT_SECRET_PATH)));

		List<String> scopes = Collections.singletonList("https://www.googleapis.com/auth/drive.readonly");

		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(httpTransport, JSON_FACTORY,
				clientSecrets, scopes).setDataStoreFactory(new FileDataStoreFactory(new File(TOKENS_DIRECTORY_PATH)))
				.setAccessType("offline").build();

		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
	}

	private static void downloadFilesFromFolder(HttpTransport httpTransport, Credential credential, String folderId,
			String destinationFolder) throws IOException {
		Drive driveService = new Drive.Builder(httpTransport, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME).build();

		// Get the list of files in the folder
		List<com.google.api.services.drive.model.File> files = driveService.files().list()
				.setQ("'" + folderId + "' in parents").execute().getFiles();

		if (files != null && !files.isEmpty()) {
			for (com.google.api.services.drive.model.File file : files) {
				downloadFile(driveService, file.getId(), destinationFolder);
			}
		} else {
			System.out.println("No files found in the specified folder.");
		}
	}

	private static void downloadFile(Drive driveService, String fileId, String destinationFolder) throws IOException {
		com.google.api.services.drive.model.File file = driveService.files().get(fileId).execute();
		String originalFileName = file.getName();

		// Sanitize the file name to remove illegal characters
		String sanitizedFileName = originalFileName.replaceAll("[<>:\"/\\\\|?*]", "_");

		String destinationFileName = destinationFolder + File.separator + sanitizedFileName;

		Path destinationFilePath = Paths.get(destinationFileName);
		Files.createDirectories(destinationFilePath.getParent());

		OutputStream outputStream = new FileOutputStream(destinationFileName);
		driveService.files().get(fileId).executeMediaAndDownloadTo(outputStream);
		outputStream.close();

		System.out.println("File downloaded successfully: " + destinationFileName);
	}
}
