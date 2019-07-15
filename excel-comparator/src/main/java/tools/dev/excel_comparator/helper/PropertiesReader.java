package tools.dev.excel_comparator.helper;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import tools.dev.excel_comparator.utils.ConfigReader;

public class PropertiesReader {

	private PropertiesReader() {

	}

	public static String getSourceFileLocation() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.source.location");
	}

	public static String getFirstSourceFileName() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.source.first.filename");
	}

	public static String getSecondSourceFileName() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.source.second.filename");
	}

	public static String getReportFileName() {
		return MessageFormat.format(
				ConfigReader.getSystemConfiguration().getString("excel-comparator.source.report.filename"),
				getSecondSourceFileName());
	}

	public static List<String> getSkipSheetNames() {
		final List<String> skipSheetNames = new ArrayList<>();
		final String skipSheets = ConfigReader.getSystemConfiguration()
				.getString("excel-comparator.reader.skip.sheet.names");
		if (skipSheets.contains(",")) {
			skipSheetNames.addAll(Arrays.asList(skipSheets.split(",")));
		} else {
			skipSheetNames.add(skipSheets);
		}
		return skipSheetNames;
	}

	public static List<String> getReplaceAllPatterns() {
		final List<String> replaceAllPatterns = new ArrayList<>();
		final String replacePatterns = ConfigReader.getSystemConfiguration()
				.getString("excel-comparator.reader.replace.patterns.all");
		if (replacePatterns.contains(",")) {
			replaceAllPatterns.addAll(Arrays.asList(replacePatterns.split(",")));
		} else {
			replaceAllPatterns.add(replacePatterns);
		}
		return replaceAllPatterns;
	}

	public static List<String> getReplacePatterns() {
		final List<String> replacePatterns = new ArrayList<>();
		final String allReplacePatterns = ConfigReader.getSystemConfiguration()
				.getString("excel-comparator.reader.replace.patterns");
		if (allReplacePatterns.contains(",")) {
			replacePatterns.addAll(Arrays.asList(allReplacePatterns.split(",")));
		} else {
			replacePatterns.add(allReplacePatterns);
		}
		return replacePatterns;
	}

	public static int getFirstKeyColumnIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.reader.firt.key.column.index");
	}

	public static int getSecondKeyColumnIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.reader.second.key.column.index");
	}

	public static int getValueColumnIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.reader.value.column.index");
	}

	public static int getStyleSheetIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.style.reader.style.sheet.index");
	}

	public static int getSummarySheetIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.style.reader.summary.sheet.index");
	}

	public static int getHeaderRowIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.style.reader.header.row.index");
	}

	public static int getFirstHeaderCellIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.style.reader.first.header.cell.index");
	}

	public static List<String> getReportRowHeaders() {
		final List<String> reportRowHeaders = new ArrayList<>();
		String allReplacePatterns = ConfigReader.getSystemConfiguration()
				.getString("excel-comparator.style.writer.report.row.headers1");
		if (allReplacePatterns.contains(",")) {
			reportRowHeaders.addAll(Arrays.asList(allReplacePatterns.split(",")));
		} else {
			reportRowHeaders.add(allReplacePatterns);
		}

		allReplacePatterns = ConfigReader.getSystemConfiguration()
				.getString("excel-comparator.style.writer.report.row.headers2");
		if (allReplacePatterns.contains(",")) {
			reportRowHeaders.addAll(Arrays.asList(allReplacePatterns.split(",")));
		} else {
			reportRowHeaders.add(allReplacePatterns);
		}
		return reportRowHeaders;
	}

	public static int getReportRowHeaderStartIndex() {
		return ConfigReader.getSystemConfiguration()
				.getInt("excel-comparator.style.writer.report.row.header.start.index");
	}

	public static int getReportRowHeaderColStartIndex() {
		return ConfigReader.getSystemConfiguration()
				.getInt("excel-comparator.style.writer.report.row.header.col.start.index");
	}

	public static int getReportRowStartIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.style.writer.report.row.start.index");
	}

	public static int getReportColStartIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.style.writer.report.col.start.index");
	}

	public static String getSummarySheetName() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.summary.sheet.name");
	}

	public static List<String> getSummarySheetHeaders() {
		final List<String> summarySheetHeaders = new ArrayList<>();
		final String headers = ConfigReader.getSystemConfiguration()
				.getString("excel-comparator.summary.sheet.headers");
		if (headers.contains(",")) {
			summarySheetHeaders.addAll(Arrays.asList(headers.split(",")));
		} else {
			summarySheetHeaders.add(headers);
		}
		return summarySheetHeaders;
	}

	public static String getSummarySheetTotalKeyword() {
		return ConfigReader.getSystemConfiguration().getString("excel-comparator.summary.sheet.total.keyword");
	}

	public static int getSummaryRowHeaderRowStartIndex() {
		return ConfigReader.getSystemConfiguration()
				.getInt("excel-comparator.summary.sheet.row.header.row.start.index");
	}

	public static int getSummaryRowHeaderColStartIndex() {
		return ConfigReader.getSystemConfiguration()
				.getInt("excel-comparator.summary.sheet.row.header.col.start.index");
	}

	public static int getSummaryRowStartIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.summary.sheet.row.start.index");
	}

	public static int getSummaryColStartIndex() {
		return ConfigReader.getSystemConfiguration().getInt("excel-comparator.summary.sheet.col.start.index");
	}
}
