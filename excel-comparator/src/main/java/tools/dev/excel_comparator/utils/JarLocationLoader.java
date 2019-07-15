package tools.dev.excel_comparator.utils;

import java.io.File;

import org.apache.commons.lang3.StringUtils;

public class JarLocationLoader {
	public static String jarFileLocation;

	private JarLocationLoader() {

	}

	private static void populateJARFileLocation() {
		jarFileLocation = new File(ClassLoader.getSystemClassLoader().getResource(".").getPath()).getAbsolutePath();
		if (jarFileLocation.contains("\\")) {
			// Comment this line when building
			jarFileLocation = jarFileLocation.substring(0, jarFileLocation.lastIndexOf("\\"));
			jarFileLocation = jarFileLocation.replace("\\", "/").replaceFirst("/", "//");
		}
	}

	public static String getJARFileLocation() {
		if (StringUtils.isBlank(jarFileLocation)) {
			populateJARFileLocation();
		}
		return jarFileLocation;
	}
}
