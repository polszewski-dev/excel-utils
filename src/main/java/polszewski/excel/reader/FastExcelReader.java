package polszewski.excel.reader;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class FastExcelReader implements ExcelReader {
	private static class RowData {
		private final List<Object> cellValues = new ArrayList<>();

		void add(int colIdx, Object cellValue) {
			resize(cellValues, colIdx + 1);
			cellValues.set(colIdx, cellValue);
		}

		List<Object> getCellValues() {
			return cellValues;
		}

		@Override
		public String toString() {
			return "RowData{" +
					"cellValues=" + cellValues +
					'}';
		}

		void cleanUp() {
			cellValues.clear();
		}
	}

	private static class SheetData {
		private final String sheetName;
		private final List<RowData> rows = new ArrayList<>();

		SheetData(String sheetName) {
			this.sheetName = sheetName;
		}

		void add(int rowIdx, int colIdx, Object cellValue) {
			if (rowIdx < 0 || colIdx < 0) {
				throw new RuntimeException("rowIdx < 0 || colIdx < 0");
			}
			resize(rows, rowIdx + 1);
			RowData rowData = rows.get(rowIdx);
			if (rowData == null) {
				rowData = new RowData();
				rows.set(rowIdx, rowData);
			}
			rowData.add(colIdx, cellValue);
		}

		String getSheetName() {
			return sheetName;
		}

		List<RowData> getRows() {
			return rows;
		}

		@Override
		public String toString() {
			return "SheetData{" +
					"rows=" + rows +
					'}';
		}

		void cleanUp() {
			for (RowData row : rows) {
				if (row != null) {
					row.cleanUp();
				}
			}
			rows.clear();
		}
	}

	private final File file;
	private XMLReader parser;
	private Iterator<InputStream> sheets;
	private SheetData currentSheet;
	private int currentRowIdx;
	private RowData currentRow;
	private int currentColIdx;
	private Object currentCell;

	public FastExcelReader(File file) {
		this.file = file;
	}

	@Override
	public void close() {
		parser = null;
		sheets = null;
		if (currentSheet != null) {
			currentSheet.cleanUp();
			currentSheet = null;
		}
		currentRow = null;
		currentCell = null;
	}

	@Override
	public boolean nextSheet() {
		currentRowIdx = 0;
		currentRow = null;
		currentColIdx = 0;
		currentCell = null;

		try {
			if (sheets == null) {
				OPCPackage pkg = OPCPackage.open(file.getAbsolutePath());
				XSSFReader r = new XSSFReader(pkg);
				SharedStringsTable sst = r.getSharedStringsTable();

				parser = fetchSheetParser(sst);
				sheets = r.getSheetsData();
			}

			if (sheets.hasNext()) {
				try (InputStream sheet = sheets.next()) {
					String currentSheetName = ((XSSFReader.SheetIterator) sheets).getSheetName();
					currentSheet = new SheetData(currentSheetName);
					InputSource sheetSource = new InputSource(sheet);
					parser.parse(sheetSource);
				}
				return true;
			}
			currentSheet = null;
			return false;
		}
		catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	@Override
	public boolean nextRow() {
		currentColIdx = 0;
		currentCell = null;

		while (currentRowIdx < currentSheet.getRows().size()) {
			currentRow = currentSheet.getRows().get(currentRowIdx++);
			if (currentRow != null) {
				return true;
			}
		}
		currentRow = null;
		return false;
	}

	@Override
	public boolean nextCell() {
		if (currentColIdx < currentRow.getCellValues().size()) {
			currentCell = currentRow.getCellValues().get(currentColIdx++);
			return true;
		}
		currentCell = null;
		return false;
	}

	@Override
	public String getCurrentSheetName() {
		return currentSheet.getSheetName();
	}

	@Override
	public int getCurrentRowIdx() {
		return currentRowIdx - 1;
	}

	@Override
	public int getCurrentColIdx() {
		return currentColIdx - 1;
	}

	@Override
	public String getCellStringValue(int colNo) {
		Object value = colNo < currentRow.getCellValues().size() ? currentRow.getCellValues().get(colNo) : null;
		return cellValueToString(value);
	}

	@Override
	public String getCurrentCellStringValue() {
		return cellValueToString(currentCell);
	}

	private static String cellValueToString(Object value) {
		if (value == null) {
			return null;
		}
		return value.toString();
	}

	private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;

		private String cellName;

		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@Override
		public void startElement(String uri, String localName, String name,
								 Attributes attributes) {
			// c => cell
			if (name.equals("c")) {
				// Print the cell reference
				cellName = attributes.getValue("r");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				nextIsString = cellType != null && cellType.equals("s");
			}
			// Clear contents cache
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v")) {
				String cellValue = lastContents;

				int rowIdx = parseRowIdx(cellName);
				int colIdx = parseColIdx(cellName);
				currentSheet.add(rowIdx - 1, colIdx - 1, cellValue);
				cellName = null;
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) {
			lastContents += new String(ch, start, length);
		}
	}

	private static int parseRowIdx(String value) {
		int result = 0;
		for (int i = findFirstDigit(value); i < value.length(); ++i) {
			result = result * 10 + value.charAt(i) - '0';
		}
		return result;
	}

	private static int parseColIdx(String value) {
		int result = 0;
		int idx = findFirstDigit(value);
		for (int i = 0; i < idx; ++i) {
			char c = value.charAt(i);
			result = result * ('Z' - 'A' + 1) + c - ('a' <= c && c <= 'z' ? 'a' : 'A') + 1;
		}
		return result;
	}

	private static int findFirstDigit(String value) {
		for (int i = 0; i < value.length(); ++i) {
			if (Character.isDigit(value.charAt(i))) {
				return i;
			}
		}
		return -1;
	}

	private static void resize(List<?> list, int size) {
		while (list.size() < size) {
			list.add(null);
		}
	}
}