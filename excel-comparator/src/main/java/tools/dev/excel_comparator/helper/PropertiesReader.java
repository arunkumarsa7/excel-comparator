package tools.dev.excel_comparator.helper;

import java.text.MessageFormat;

import tools.dev.excel_comparator.utils.ConfigReader;

public class PropertiesReader {

	private PropertiesReader() {

	}

	public static String getSourceFileLocation() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.source.location");
	}

	public static String getFirstSourceFileName() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.source.first.filename");
	}

	public static String getSecondSourceFileName() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.source.second.filename");
	}

	public static String getReportFileName() {
		return MessageFormat.format(
				ConfigReader.getSystemConfiguration().getString("excel-comparator.source.report.filename"),
				getSecondSourceFileName());
	}
}
