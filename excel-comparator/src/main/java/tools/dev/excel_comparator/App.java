package tools.dev.excel_comparator;

import tools.dev.excel_comparator.utils.ExcelWriter;

/**
 *
 * @author Arun Kumar S A
 *
 */
public class App {
	public static void main(final String[] args) {
		final ExcelWriter excelWriter = new ExcelWriter();
		excelWriter.createWorkBook();
	}
}
