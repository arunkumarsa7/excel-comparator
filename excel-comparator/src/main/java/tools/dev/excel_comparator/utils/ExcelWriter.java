package tools.dev.excel_comparator.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import tools.dev.excel_comparator.helper.PropertiesReader;
import tools.dev.excel_comparator.service.CompareXSSFWorkbookHelper;

public class ExcelWriter {

	public void createWorkBook() {
		try (final Workbook firstWorkbook = ExcelReader.readExcelFile(PropertiesReader.getSourceFileLocation(),
				PropertiesReader.getFirstSourceFileName());
				final Workbook secondWorkbook = ExcelReader.readExcelFile(PropertiesReader.getSourceFileLocation(),
						PropertiesReader.getSecondSourceFileName());
				final Workbook workbook = new XSSFWorkbook();
				final FileOutputStream fileOutputStream = new FileOutputStream(PropertiesReader.getSourceFileLocation()
						+ File.separator + PropertiesReader.getReportFileName());) {
			final Map<String, Map<String, Map<String, String>>> firstWorkbookData = ExcelReader
					.convertWorkbookToMap(firstWorkbook);
			final Map<String, Map<String, Map<String, String>>> secondWorkbookData = ExcelReader
					.convertWorkbookToMap(secondWorkbook);
			final Map<String, Map<String, Map<String, String>>> newWorkbookData = CompareXSSFWorkbookHelper
					.compareXSSFWorkbooks(firstWorkbookData, secondWorkbookData);
			ExcelStyleReader.initializeStyleReader(workbook);
			populateDataToWorkbook(workbook, newWorkbookData);
			workbook.write(fileOutputStream);
		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

	private static void populateDataToWorkbook(final Workbook workbook,
			final Map<String, Map<String, Map<String, String>>> newWorkbookData) throws IOException {
		if (!newWorkbookData.isEmpty()) {
			ExcelStyleReader.populateStyleReaderSheet();
			SummarySheetWriter.createSummarySheet(workbook, newWorkbookData);
			for (final Entry<String, Map<String, Map<String, String>>> entry : newWorkbookData.entrySet()) {
				int rowStartIndex = PropertiesReader.getReportRowStartIndex();
				final String sheetName = entry.getKey();
				final Sheet sheet = workbook.createSheet(sheetName);
				writeSheetHeaders(sheet);
				final Map<String, Map<String, String>> firstWorkbookSheetContent = entry.getValue();
				if (!firstWorkbookSheetContent.isEmpty()) {
					for (final Entry<String, Map<String, String>> cellEntry : firstWorkbookSheetContent.entrySet()) {
						int colStartIndex = PropertiesReader.getReportColStartIndex();
						final Map<String, String> cellValue = cellEntry.getValue();
						final Map.Entry<String, String> cellValueEntry = cellValue.entrySet().iterator().next();
						final Row row = sheet.createRow(rowStartIndex++);
						final Cell cell1 = row.createCell(colStartIndex++);
						cell1.setCellValue(cellValueEntry.getKey());
						final Cell cell2 = row.createCell(colStartIndex);
						cell2.setCellValue(cellValueEntry.getValue());
						cell2.setCellStyle(ExcelStyleReader.getWrapTextCellStyle());
					}
				}
			}
		}
	}

	private static void writeSheetHeaders(final Sheet sheet) {
		int rowHeaderColStartIndex = PropertiesReader.getReportRowHeaderStartIndex();
		final Row row = sheet.createRow(PropertiesReader.getReportRowHeaderColStartIndex());
		for (final String rowHeader : PropertiesReader.getReportRowHeaders()) {
			final Cell cell = row.createCell(rowHeaderColStartIndex++);
			if (rowHeaderColStartIndex == 1) {
				cell.setCellStyle(ExcelStyleReader.getFirstHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1, ExcelStyleReader.getFirstHeaderCellWidth());
			} else if (rowHeaderColStartIndex == 2) {
				cell.setCellStyle(ExcelStyleReader.getSecondHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1, ExcelStyleReader.getSecondHeaderCellWidth());
			} else {
				cell.setCellStyle(ExcelStyleReader.getThirdHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1, ExcelStyleReader.getThirdHeaderCellWidth());
			}
			cell.setCellValue(rowHeader);
		}
	}

}
