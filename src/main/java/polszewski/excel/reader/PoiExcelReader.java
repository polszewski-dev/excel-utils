package polszewski.excel.reader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

/**
 * User: POlszewski
 * Date: 2015-06-11
 */
public class PoiExcelReader implements ExcelReader {
	private final Workbook workbook;
	private int sheetNo;
	private Sheet currentSheet;
	private Iterator<Row> rowIterator;
	private Row currentRow;
	private Iterator<Cell> cellIterator;
	private Cell currentCell;

	public PoiExcelReader(Workbook workbook) {
		this.workbook = workbook;
	}

	public PoiExcelReader(File file) throws IOException, InvalidFormatException {
		this(WorkbookFactory.create(file));
	}

	public PoiExcelReader(InputStream inputStream) throws IOException, InvalidFormatException {
		this(WorkbookFactory.create(inputStream));
	}

	@Override
	public void close() throws IOException {
		workbook.close();
	}

	@Override
	public boolean nextSheet() {
		rowIterator = null;
		currentRow = null;
		cellIterator = null;
		currentCell = null;

		if (sheetNo < workbook.getNumberOfSheets()) {
			currentSheet = workbook.getSheetAt(sheetNo);
			++sheetNo;
			return true;
		}
		return false;
	}

	@Override
	public boolean nextRow() {
		cellIterator = null;
		currentCell = null;

		if (rowIterator == null) {
			rowIterator = currentSheet.rowIterator();
		}
		if (rowIterator.hasNext()) {
			currentRow = rowIterator.next();
			return true;
		}
		currentRow = null;
		return false;
	}

	@Override
	public boolean nextCell() {
		if (cellIterator == null) {
			cellIterator = currentRow.cellIterator();
		}
		if (cellIterator.hasNext()) {
			currentCell = cellIterator.next();
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
		return currentRow.getRowNum();
	}

	@Override
	public int getCurrentColIdx() {
		return currentCell.getColumnIndex();
	}

	@Override
	public String getCellStringValue(int colNo) {
		return getStringCellValue(currentRow.getCell(colNo));
	}

	@Override
	public String getCurrentCellStringValue() {
		return getStringCellValue(currentCell);
	}

	private static String getStringCellValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				return numericCellToString(cell);
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue().trim();
			case Cell.CELL_TYPE_BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case Cell.CELL_TYPE_BLANK:
				return null;
			case Cell.CELL_TYPE_FORMULA:
				return formulaCellToString(cell);
			default:
				return null;
		}
	}

	private static String numericCellToString(Cell cell) {
		double value = cell.getNumericCellValue();
		return value % 1 == 0 ? String.valueOf((long)value) : String.valueOf(value);
	}

	private static String formulaCellToString(Cell cell) {
		return cell.getCellFormula();
	}
}
