pulic class SummarySheetWriter {
private static String summarySheetName ="SUMMARY";
private static int rowHeaderStartIndex = 0;
private static int rowHeaderColStartIndex = 0;
  
static String[] summarySheetHeaders = { "Attribute", "Reference Count" };
      
private SummarySheetWriter(){
}

public static void createSummarySheet(final Workbook workbook,final Map<String, Map<String, Map<String, String>>> newWorkbookData){
final Sheet sheet = workbook.createSheet(summarySheetName);
writeSheetHeaders(sheet);
for (final Entry<String, Map<String, Map<String, String>>> entry : newWorkbookData.entrySet()) {
        final String sheetName = entry.getKey();
				final Map<String, Map<String, String>> firstWorkbookSheetContent = entry.getValue();
        if (!firstWorkbookSheetContent.isEmpty()) {
					for (final Entry<String, Map<String, String>> cellEntry : firstWorkbookSheetContent.entrySet()) {
           int colStartIndex = 0;
          }
      }
   }
}

private static void writeSheetHeaders(final Sheet sheet) {
		final Row row = sheet.createRow(rowHeaderStartIndex);
		for (final String rowHeader : Arrays.asList(summarySheetHeaders)) {
			final Cell cell = row.createCell(rowHeaderColStartIndex++);
			if (rowHeaderColStartIndex == 1) {
				cell.setCellStyle(ExcelStyleReader.getSummarySheetFirstHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1, ExcelStyleReader.getSummarySheetFirstHeaderCellWidth());
			} else if (rowHeaderColStartIndex == 2) {
				cell.setCellStyle(ExcelStyleReader.getSummarySheetSecondHeaderCellStyle());
				sheet.setColumnWidth(rowHeaderColStartIndex - 1, ExcelStyleReader.getSummarySheetSecondHeaderCellWidth());
			}
			cell.setCellValue(rowHeader);
		}
	}

}
