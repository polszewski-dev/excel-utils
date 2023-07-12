package polszewski.excel.writer.style;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public record Font(
		String name,
		Integer size,
		Color color,
		Boolean bold,
		Boolean italic
) {
	public static final Font DEFAULT = new Font(null, null, null, null, null);

	public Font name(String name) {
		return new Font(name, size, color, bold, italic);
	}

	public Font size(Integer size) {
		return new Font(name, size, color, bold, italic);
	}

	public Font color(Color color) {
		return new Font(name, size, color, bold, italic);
	}

	public Font bold(boolean bold) {
		return new Font(name, size, color, bold, italic);
	}

	public Font italic(boolean italic) {
		return new Font(name, size, color, bold, italic);
	}

	public Font normalFont() {
		return new Font(name, size, color, false, false);
	}
}
