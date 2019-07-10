package tools.dev.excel_comparator.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	public static final DataFormatter formatter = new DataFormatter();

	private ExcelReader() {

	}

	public static Workbook readExcelFile(final String filePath, final String fileName) throws IOException {
		return new XSSFWorkbook(new FileInputStream(new File(filePath + File.separator + fileName)));
	}

	public static Map<String, String> covertSheetContentToMap(final Workbook workbook, final int sheetIndex) {
		final Map<String, String> sheetContentMap = new HashMap<>();
		final Sheet sheet = workbook.getSheetAt(sheetIndex);
		final int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for (int i = 0; i < rowCount + 1; i++) {
			final Row row = sheet.getRow(i);
			sheetContentMap.put(getSheetContentToMapKey(row), getSheetContentToMapValue(row));
		}
		return sheetContentMap;
	}

	private static String getSheetContentToMapKey(final Row row) {
		return formatter.formatCellValue(row.getCell(0)) + formatter.formatCellValue(row.getCell(1));
	}

	private static String getSheetContentToMapValue(final Row row) {
		return formatter.formatCellValue(row.getCell(1));
	}
}
