package usecases_related_chrome;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ConfigReader {
    private Properties properties;

    public ConfigReader(String configFilePath) throws IOException {
        properties = new Properties();
        File configFile = new File(configFilePath);

        try (InputStream input = new FileInputStream(configFile)) {
            properties.load(input);
        }
    }

    public Properties getProperties() {
        return properties;
    }

    public String getClientSecretsPath() {
        return properties.getProperty("client_secrets_path");
    }

    public String getTokensDirectoryPath() {
        return properties.getProperty("tokens_directory_path");
    }

    public String getFolderId() {
        return properties.getProperty("folder_id");
    }

    public String getDestinationFolderPath() {
        return properties.getProperty("destination_folder_path");
    }
}
