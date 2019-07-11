package tools.dev.excel_comparator.service;

import java.io.IOException;
import java.util.Map;

import tools.dev.excel_comparator.utils.ExcelReader;

public class CompareXSSFWorkbookHelperTest {
	static String fileLocation;
	static String firstFileName;
	static String secondFileName;
	static int sheetNumber;

	static void setup() {
		fileLocation = "C:\\Users\\Arun Kumar S A\\Desktop";
		firstFileName = "testexcel.xlsx";
		secondFileName = "testexcel - Copy.xlsx";
		sheetNumber = 0;
	}

	public static void main(final String[] args) {
		setup();
		try {
			final Map<String, Map<String, String>> firstFile = ExcelReader.covertSheetContentToMap(
					ExcelReader.readExcelFile(fileLocation, firstFileName).getSheetAt(sheetNumber));
			final Map<String, Map<String, String>> secondFile = ExcelReader.covertSheetContentToMap(
					ExcelReader.readExcelFile(fileLocation, secondFileName).getSheetAt(sheetNumber));
			final Map<String, Map<String, String>> differenceMap = CompareXSSFWorkbookHelper
					.compareXSSFWorkbookSheets(firstFile, secondFile);
			if (differenceMap.isEmpty()) {
				System.out.println("Both files are equal");
			} else {
				System.out.println("Both file are different");
				System.out.println("Differences are:");
				for (final Map.Entry<String, Map<String, String>> entry : differenceMap.entrySet()) {
					System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());
				}
			}
		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

}
