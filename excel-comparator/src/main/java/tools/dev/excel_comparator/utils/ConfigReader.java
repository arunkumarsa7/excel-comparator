package tools.dev.excel_comparator.utils;

import java.util.ResourceBundle;

public class ConfigReader {
	private static ResourceBundle configuration;

	private ConfigReader() {

	}

	private static void readSystemConfigFile() {
		configuration = ResourceBundle.getBundle("config");
	}

	public static ResourceBundle getSystemConfiguration() {
		if (configuration == null) {
			readSystemConfigFile();
		}
		return configuration;
	}
}
