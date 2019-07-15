package tools.dev.excel_comparator.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyleReader {
	static String fileLocation = "C:\\Users\\Arun Kumar S A\\Desktop";
	static String fileName = "Attribute_References_19.0.xlsx";
	static Sheet styleReaderSheet;
	static Sheet summaryStyleReaderSheet;
	static int styleSheetIndex = 1;
	static int summarySheetIndex = 0;
	static int headerRowIndex = 0;
	static int firstHeaderCellIndex = 0;
	private static Workbook parentWorkbook;

	private ExcelStyleReader() {

	}

	public static void populateStyleReaderSheet() throws IOException {
		try (final FileInputStream fis = new FileInputStream(new File(fileLocation + File.separator + fileName));
				final Workbook workbook = new XSSFWorkbook(fis);) {
			styleReaderSheet = workbook.getSheetAt(styleSheetIndex);
			summaryStyleReaderSheet = workbook.getSheetAt(summarySheetIndex);
		}
	}

	public static CellStyle getFirstHeaderCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle
				.cloneStyleFrom(styleReaderSheet.getRow(headerRowIndex).getCell(firstHeaderCellIndex).getCellStyle());
		return workbookCellStyle;
	}

	public static CellStyle getSecondHeaderCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle.cloneStyleFrom(
				styleReaderSheet.getRow(headerRowIndex).getCell(firstHeaderCellIndex + 1).getCellStyle());
		return workbookCellStyle;
	}

	public static CellStyle getThirdHeaderCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle.cloneStyleFrom(
				styleReaderSheet.getRow(headerRowIndex).getCell(firstHeaderCellIndex + 2).getCellStyle());
		return workbookCellStyle;
	}

	public static CellStyle getWrapTextCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle.setWrapText(true);
		return workbookCellStyle;
	}

	public static int getFirstHeaderCellWidth() {
		return styleReaderSheet.getColumnWidth(firstHeaderCellIndex);
	}

	public static int getSecondHeaderCellWidth() {
		return styleReaderSheet.getColumnWidth(firstHeaderCellIndex + 1);
	}

	public static int getThirdHeaderCellWidth() {
		return styleReaderSheet.getColumnWidth(firstHeaderCellIndex + 2);
	}

	public static void initializeStyleReader(final Workbook workbook) {
		parentWorkbook = workbook;
	}

	public static CellStyle getSummarySheetFirstHeaderCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle.cloneStyleFrom(
				summaryStyleReaderSheet.getRow(headerRowIndex).getCell(firstHeaderCellIndex).getCellStyle());
		return workbookCellStyle;
	}

	public static int getSummarySheetFirstHeaderCellWidth() {
		return summaryStyleReaderSheet.getColumnWidth(firstHeaderCellIndex);
	}

	public static CellStyle getSummarySheetSecondHeaderCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle.cloneStyleFrom(
				summaryStyleReaderSheet.getRow(headerRowIndex).getCell(firstHeaderCellIndex + 1).getCellStyle());
		return workbookCellStyle;
	}

	public static int getSummarySheetSecondHeaderCellWidth() {
		return summaryStyleReaderSheet.getColumnWidth(firstHeaderCellIndex + 1);
	}

	public static CellStyle getSummarySheetNormalCellStyle() {
		final CellStyle workbookCellStyle = parentWorkbook.createCellStyle();
		workbookCellStyle.cloneStyleFrom(
				summaryStyleReaderSheet.getRow(headerRowIndex + 1).getCell(firstHeaderCellIndex).getCellStyle());
		return workbookCellStyle;
	}
}
