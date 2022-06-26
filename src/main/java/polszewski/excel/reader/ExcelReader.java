package polszewski.excel.reader;

import java.io.Closeable;
import java.io.IOException;

/**
 * User: POlszewski
 * Date: 2015-06-11
 */
public interface ExcelReader extends Closeable {
	void close() throws IOException;

	boolean nextSheet();

	boolean nextRow();

	boolean nextCell();

	String getCurrentSheetName();

	int getCurrentRowIdx();

	int getCurrentColIdx();

	String getCellStringValue(int colNo);

	String getCurrentCellStringValue();
}
