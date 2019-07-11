package tools.dev.excel_comparator.service;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

/**
 *
 * @author Arun Kumar S A
 *
 */
public class CompareXSSFWorkbookHelper {
	private CompareXSSFWorkbookHelper() {

	}

	public static Map<String, Map<String, String>> compareXSSFWorkbookSheets(
			final Map<String, Map<String, String>> firstWorkbookSheet,
			final Map<String, Map<String, String>> secondWorkbookSheet) {
		final Map<String, Map<String, String>> differenceMap = new LinkedHashMap<>();
		for (final Entry<String, Map<String, String>> entry : firstWorkbookSheet.entrySet()) {
			final String key = entry.getKey();
			if (secondWorkbookSheet.containsKey(key)) {
				secondWorkbookSheet.remove(key);
			}
		}
		if (!secondWorkbookSheet.isEmpty()) {
			for (final Entry<String, Map<String, String>> entry : secondWorkbookSheet.entrySet()) {
				differenceMap.put(entry.getKey(), entry.getValue());
			}
		}
		return differenceMap;
	}

	public static Map<String, Map<String, Map<String, String>>> compareXSSFWorkbooks(
			final Map<String, Map<String, Map<String, String>>> firstWorkbook,
			final Map<String, Map<String, Map<String, String>>> secondWorkbook) {
		final Map<String, Map<String, Map<String, String>>> differenceMap = new LinkedHashMap<>();
		for (final Entry<String, Map<String, Map<String, String>>> entry : firstWorkbook.entrySet()) {
			final String sheetName = entry.getKey();
			final Map<String, Map<String, String>> firstWorkbookSheetContent = entry.getValue();
			if (secondWorkbook.containsKey(sheetName)) {
				final Map<String, Map<String, String>> sheetDifferenceMap = compareXSSFWorkbookSheets(
						firstWorkbookSheetContent, secondWorkbook.get(sheetName));
				if (!sheetDifferenceMap.isEmpty()) {
					differenceMap.put(sheetName, sheetDifferenceMap);
				} else {
					differenceMap.put(sheetName, new LinkedHashMap<String, Map<String, String>>());
				}
				secondWorkbook.remove(sheetName);
			}
		}
		if (!secondWorkbook.isEmpty()) {
			for (final Entry<String, Map<String, Map<String, String>>> entry : secondWorkbook.entrySet()) {
				differenceMap.put(entry.getKey(), entry.getValue());
			}
		}
		return differenceMap;
	}

}
