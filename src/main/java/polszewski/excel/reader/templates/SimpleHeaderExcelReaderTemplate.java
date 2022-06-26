package polszewski.excel.reader.templates;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import polszewski.excel.reader.ExcelReader;
import polszewski.excel.reader.PoiExcelReader;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.function.Function;
import java.util.regex.Pattern;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public abstract class SimpleHeaderExcelReaderTemplate<T> {
	private ExcelReader excelReader;
	private Map<String, Integer> header;

	public List<T> read(String filePath) {
		try {
			return readHelper(new PoiExcelReader(new File(filePath)));
		} catch (IOException | InvalidFormatException e) {
			throw new ExcelException(e);
		}
	}

	public List<T> read(InputStream inputStream) {
		try {
			return readHelper(new PoiExcelReader(inputStream));
		} catch (IOException | InvalidFormatException e) {
			throw new ExcelException(e);
		}
	}

	public List<T> readFromResource(String path) {
		return read(getClass().getResourceAsStream(path));
	}

	private List<T> readHelper(PoiExcelReader excelReader1) throws IOException {
		try (ExcelReader excelReader = excelReader1) {
			this.excelReader = excelReader;

			List<T> result = new ArrayList<>();

			while (excelReader.nextSheet()) {
				onNewSheet();

				if (!excelReader.nextRow()) {
					onEmptySheet();
				} else {
					header = parseHeader();

					while (excelReader.nextRow()) {
						T record = readRecord(excelReader);

						if (record != null) {
							result.add(record);
						}
					}
				}
			}

			return result;
		} finally {
			this.excelReader = null;
		}
	}

	protected void onNewSheet() {}

	protected void onEmptySheet() {}

	protected abstract T readRecord(ExcelReader excelReader);

	protected Map<String, Integer> parseHeader() {
		Map<String, Integer> result = new TreeMap<>(String.CASE_INSENSITIVE_ORDER);

		while (excelReader.nextCell()) {
			String value = excelReader.getCurrentCellStringValue();
			if (value != null) {
				result.put(value.trim(), excelReader.getCurrentColIdx());
			}
		}
		return result;
	}

	protected String getCurrentSheetName() {
		return excelReader.getCurrentSheetName();
	}

	protected String getString(String colName) {
		Integer colNo = header.get(colName);
		if (colNo == null) {
			return null;
		}
		return excelReader.getCellStringValue(colNo);
	}

	protected String getString(int colNo) {
		return excelReader.getCellStringValue(colNo);
	}

	public Boolean getBoolean(String name) {
		String value = getString(name);
		return value != null && !value.isEmpty() ? Boolean.parseBoolean(value) : null;
	}

	public boolean getBooleanFalseOnNull(String name) {
		Boolean result = getBoolean(name);
		return result != null ? result : false;
	}

	public Integer getInteger(String name) {
		String value = getString(name);
		return value != null && !value.isEmpty() ? Integer.parseInt(value) : null;
	}

	public Double getDouble(String name) {
		String value = getString(name);
		return value != null && !value.isEmpty() ? Double.parseDouble(value) : null;
	}

	public List<String> getStringList(String name, String separator) {
		String value = getString(name);
		if (value == null || value.isEmpty()) {
			return Collections.emptyList();
		}
		return Arrays.asList(value.split(Pattern.quote(separator)));
	}

	public <T> List<T> getList(String name, String separator, Function<String, T> converter) {
		List<String> stringList = getStringList(name, separator);
		if (stringList.isEmpty()) {
			return Collections.emptyList();
		}
		List<T> result = new ArrayList<>(stringList.size());
		for (String s : stringList) {
			result.add(converter.apply(s));
		}
		return result;
	}

	public List<String> getFieldsStartingWith(String prefix) {
		List<String> result = new ArrayList<>();

		for (String fieldName : header.keySet()) {
			if (fieldName.startsWith(prefix)) {
				result.add(fieldName);
			}
		}
		return result;
	}
}
