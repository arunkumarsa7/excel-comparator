package tools.dev.excel_comparator.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	static int firtKeyColumn = 0;
	static int secondKeyColumn = 1;
	static int valueColumn = 1;
	static String skipSheetName = "Summary";

	public static final DataFormatter formatter = new DataFormatter();

	private ExcelReader() {

	}

	public static Workbook readExcelFile(final String filePath, final String fileName) throws IOException {
		try (final FileInputStream fis = new FileInputStream(new File(filePath + File.separator + fileName));
				final Workbook workbook = new XSSFWorkbook(fis);) {
			return workbook;
		}
	}

	public static Map<String, Map<String, Map<String, String>>> convertWorkbookToMap(final Workbook workbook) {
		final Map<String, Map<String, Map<String, String>>> workbookMap = new LinkedHashMap<>();
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			final Sheet sheet = workbook.getSheetAt(i);
			if (!skipSheetName.equalsIgnoreCase(sheet.getSheetName())) {
				workbookMap.put(sheet.getSheetName().toUpperCase(), covertSheetContentToMap(sheet));
			}
		}
		return workbookMap;
	}

	public static Map<String, Map<String, String>> covertSheetContentToMap(final Sheet sheet) {
		final Map<String, Map<String, String>> sheetContentMap = new LinkedHashMap<>();
		if (!isSheetEmpty(sheet)) {
			final int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
			for (int i = 0; i < rowCount + 1; i++) {
				final Row row = sheet.getRow(i);
				sheetContentMap.put(getSheetContentToMapKey(row), getSheetContentToMapValue(row));
			}
		}
		return sheetContentMap;
	}

	private static String getSheetContentToMapKey(final Row row) {
		return formatCellValue(row.getCell(firtKeyColumn)) + formatCellValue(row.getCell(secondKeyColumn));
	}

	private static Map<String, String> getSheetContentToMapValue(final Row row) {
		final Map<String, String> sheetContentMap = new LinkedHashMap<>();
		sheetContentMap.put(formatter.formatCellValue(row.getCell(firtKeyColumn)),
				formatter.formatCellValue((row.getCell(valueColumn))));
		return sheetContentMap;
	}

	private static String formatCellValue(final Cell columnValue) {
		return formatter.formatCellValue(columnValue).replaceAll("\\s+", "").replace("\r\n", " ").replace("\n", " ");
	}

	private static boolean isSheetEmpty(final Sheet sheet) {
		if (sheet == null) {
			return true;
		}
		final Row row = sheet.getRow(0);
		if (row == null) {
			return true;
		}
		return StringUtils.isBlank(formatCellValue(row.getCell(0)));
	}
}
