package polszewski.excel.reader.templates;

import polszewski.excel.reader.ExcelReader;
import polszewski.excel.reader.PoiExcelReader;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.stream.Stream;

/**
 * User: POlszewski
 * Date: 2022-10-13
 */
public abstract class ExcelParser {
	protected ExcelReader excelReader;
	protected Map<String, Integer> header;

	public final void readFromXls() throws IOException {
		try (var newExcelReader = new PoiExcelReader(getExcelInputStream())) {
			this.excelReader = newExcelReader;
			iterateSheetsAccordingToParserOrder();
		} finally {
			this.excelReader = null;
			this.header = null;
		}
	}

	private void iterateSheetsAccordingToParserOrder() {
		var parserToSheets = getParserToSheetMap();

		for (var entry : parserToSheets.entrySet()) {
			var sheetParser = entry.getKey();
			var sheetNames = entry.getValue();

			for (String sheetName : sheetNames) {
				excelReader.goToSheet(sheetName);
				readSheet(sheetParser);
			}
		}
	}

	private void readSheet(ExcelSheetParser sheetParser) {
		if (excelReader.nextRow()) {
			this.header = getHeader();
			sheetParser.init(excelReader, header);
			sheetParser.readSheet();
		}
	}

	private Map<ExcelSheetParser, List<String>> getParserToSheetMap() {
		var sheetNames = new LinkedHashSet<>(excelReader.getSheetNames());
		var sheetParsers = getSheetParsers().toList();
		var parserToSheets = new LinkedHashMap<ExcelSheetParser, List<String>>();

		for (var sheetParser : sheetParsers) {
			for (var iterator = sheetNames.iterator(); iterator.hasNext(); ) {
				String sheetName = iterator.next();
				if (sheetParser.matchesSheetName(sheetName)) {
					parserToSheets.computeIfAbsent(sheetParser, x -> new ArrayList<>()).add(sheetName);
					iterator.remove();
				}
			}
		}

		assertAllSheetsHaveParser(sheetNames);
		assertAllParserHaveSheets(sheetParsers, parserToSheets);

		return parserToSheets;
	}

	private void assertAllParserHaveSheets(List<ExcelSheetParser> sheetParsers, Map<ExcelSheetParser, List<String>> parserToSheets) {
		if (parserToSheets.size() != sheetParsers.size()) {
			var copy = new ArrayList<>(sheetParsers);
			copy.removeAll(parserToSheets.keySet());
			throw new IllegalArgumentException("Couldn't find matching sheets for parsers: " + copy);
		}
	}

	private void assertAllSheetsHaveParser(Set<String> sheetNames) {
		if (!sheetNames.isEmpty()) {
			throw new IllegalArgumentException("Couldn't find matching parsers for sheets: " + String.join(", ", sheetNames));
		}
	}

	protected abstract InputStream getExcelInputStream() throws IOException;

	protected abstract Stream<ExcelSheetParser> getSheetParsers();

	protected InputStream fromResourcePath(String path) {
		return this.getClass().getResourceAsStream(path);
	}

	private Map<String, Integer> getHeader() {
		var result = new LinkedHashMap<String, Integer>();

		while (excelReader.nextCell()) {
			String value = excelReader.getCurrentCellStringValue();
			if (value != null) {
				result.put(value.trim(), excelReader.getCurrentColIdx());
			}
		}
		return result;
	}
}
