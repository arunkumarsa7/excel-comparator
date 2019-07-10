package tools.dev.excel_comparator.service;

import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

public class CompareXSSFWorkbookHelper {
	private CompareXSSFWorkbookHelper() {

	}

	public static Map<String, String> compareXSSFWorkbooks(final Map<String, String> firstWorkbook,
			final Map<String, String> secondWorkbook) {
		final Map<String, String> differenceMap = new HashMap<>();
		for (final Entry<String, String> entry : firstWorkbook.entrySet()) {
			final String key = entry.getKey();
			final String value = entry.getValue();
			if (secondWorkbook.containsKey(key)) {
				if (!secondWorkbook.get(key).equals(firstWorkbook.get(key))) {
					differenceMap.put(key, value);
				}
				secondWorkbook.remove(key);
			} else {
				differenceMap.put(key, value);
			}
		}
		if (!secondWorkbook.isEmpty()) {
			for (final Entry<String, String> entry : secondWorkbook.entrySet()) {
				differenceMap.put(entry.getKey(), entry.getValue());
			}
		}
		return secondWorkbook;
	}
}
