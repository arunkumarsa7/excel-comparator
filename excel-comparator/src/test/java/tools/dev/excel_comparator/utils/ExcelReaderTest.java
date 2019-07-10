package tools.dev.excel_comparator.utils;

import java.io.IOException;
import java.util.Map;

public class ExcelReaderTest {
	static String fileLocation;
	static String fileName;
	static int sheetNumber;

	static void setup() {
		fileLocation = "C:\\Users\\Arun Kumar S A\\Desktop";
		fileName = "testexcel.xlsx";
		sheetNumber = 0;
	}

	public static void main(final String[] args) {
		setup();
		try {
			for (final Map.Entry<String, String> entry : ExcelReader
					.covertSheetContentToMap(ExcelReader.readExcelFile(fileLocation, fileName), sheetNumber)
					.entrySet()) {
				System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());
			}
		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

}
