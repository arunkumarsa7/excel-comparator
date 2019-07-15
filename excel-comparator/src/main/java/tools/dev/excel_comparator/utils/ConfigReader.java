package tools.dev.excel_comparator.utils;

import java.io.File;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

import org.apache.commons.configuration2.FileBasedConfiguration;
import org.apache.commons.configuration2.PropertiesConfiguration;
import org.apache.commons.configuration2.builder.FileBasedConfigurationBuilder;
import org.apache.commons.configuration2.builder.fluent.Parameters;
import org.apache.commons.configuration2.ex.ConfigurationException;

public class ConfigReader {
	private static FileBasedConfiguration configuration;

	private ConfigReader() {

	}

	private static void readSystemConfigFile() {
		try {
			final String configLocation = URLDecoder.decode(
					JarLocationLoader.getJARFileLocation() + "/config/config.properties",
					StandardCharsets.UTF_8.name());
			final File propertiesFile = new File(configLocation);
			final FileBasedConfigurationBuilder<FileBasedConfiguration> builder = new FileBasedConfigurationBuilder<FileBasedConfiguration>(
					PropertiesConfiguration.class).configure(new Parameters().fileBased().setFile(propertiesFile));
			configuration = builder.getConfiguration();
		} catch (final ConfigurationException | UnsupportedEncodingException e) {
			System.out.println(e.getMessage());
		}
	}

	public static FileBasedConfiguration getSystemConfiguration() {
		if (configuration == null) {
			readSystemConfigFile();
		}
		return configuration;
	}
}
