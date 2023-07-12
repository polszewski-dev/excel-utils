package polszewski.excel.reader.templates;

import polszewski.excel.reader.ExcelReader;
import polszewski.excel.reader.PoiExcelReader;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;
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
			while (excelReader.nextSheet()) {
				if (excelReader.nextRow()) {
					this.header = getHeader();
					ExcelSheetParser sheetParser = getSheetParser(excelReader);
					sheetParser.init(excelReader, header);
					sheetParser.readSheet();
				}
			}
		} finally {
			this.excelReader = null;
			this.header = null;
		}
	}

	private ExcelSheetParser getSheetParser(ExcelReader excelReader) {
		return getSheetParsers()
				.filter(parser -> parser.matchesSheetName(excelReader))
				.findFirst()
				.orElseThrow(() -> new IllegalArgumentException("Unknown sheet: " + excelReader.getCurrentSheetName()));
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
