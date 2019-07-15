package tools.dev.excel_comparator.utils;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SummarySheetWriter {
	private static String summarySheetName = "SUMMARY";
	private static int rowHeaderStartIndex = 0;
	private static int rowHeaderColStartIndex = 0;
	private static int rowStartIndex = 1;
	private static int colStartIndex = 0;

	static String[] summarySheetHeaders = { "Attribute", "Reference Count" };
	static String totalKeyword = "Total";

	private SummarySheetWriter() {
	}

	public static void createSummarySheet(final Workbook workbook,
			final Map<String, Map<String, Map<String, String>>> newWorkbookData) {
		final Sheet sheet = workbook.createSheet(summarySheetName);
		writeSheetHeaders(sheet);
		int totalEntries = 0;
		for (final Entry<String, Integer> entry : getSummarySheetData(newWorkbookData).entrySet()) {
			colStartIndex = 0;
			final Integer numberOfEntries = entry.getValue();
			final Row row = sheet.createRow(rowStartIndex++);
			final Cell cell1 = row.createCell(colStartIndex++);
			final Cell cell2 = row.createCell(colStartIndex++);
			cell1.setCellValue(entry.getKey());
			cell2.setCellValue(numberOfEntries);
			totalEntries += numberOfEntries;
		}
		writeTotalEntries(sheet.createRow(rowStartIndex++), totalEntries);
	}

	private static Map<String, Integer> getSummarySheetData(
			final Map<String, Map<String, Map<String, String>>> newWorkbookData) {
		final Map<String, Integer> dataForSummarySheet = new LinkedHashMap<>();
		for (final Entry<String, Map<String, Map<String, String>>> entry : newWorkbookData.entrySet()) {
			dataForSummarySheet.put(entry.getKey(), entry.getValue().values().size());
		}
		return dataForSummarySheet;
	}

	private static void writeSheetHeaders(final Sheet sheet) {
		final Row row = sheet.createRow(rowHeaderStartIndex);
		for (final String rowHeader : Arrays.asList(summarySheetHeaders)) {
			final Cell cell = row.createCell(rowHeaderColStartIndex++);
			if (rowHeaderColStartIndex == 1) {
				cell.setCellStyle(ExcelStyleReader.getSummarySheetFirstHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1,
						ExcelStyleReader.getSummarySheetFirstHeaderCellWidth());
			} else if (rowHeaderColStartIndex == 2) {
				cell.setCellStyle(ExcelStyleReader.getSummarySheetSecondHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1,
						ExcelStyleReader.getSummarySheetSecondHeaderCellWidth());
			}
			cell.setCellValue(rowHeader);
		}
	}

	private static void writeTotalEntries(final Row row, final int totalEntries) {
		colStartIndex = 0;
		final Cell cell1 = row.createCell(colStartIndex++);
		final Cell cell2 = row.createCell(colStartIndex++);
		cell1.setCellValue(totalKeyword);
		cell2.setCellValue(totalEntries);
		cell1.setCellStyle(ExcelStyleReader.getSummarySheetNormalCellStyle());
		cell2.setCellStyle(ExcelStyleReader.getSummarySheetNormalCellStyle());
	}

}
