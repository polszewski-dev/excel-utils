package polszewski.excel.reader.templates;

/**
 * User: POlszewski
 * Date: 2019-01-13
 */
public class ExcelException extends RuntimeException {
	public ExcelException() {
	}

	public ExcelException(String message) {
		super(message);
	}

	public ExcelException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelException(Throwable cause) {
		super(cause);
	}

	public ExcelException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}
}