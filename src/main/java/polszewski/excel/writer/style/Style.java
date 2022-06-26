package polszewski.excel.writer.style;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Objects;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public class Style {
	public static final Style DEFAULT = new Style();

	private Short alignment;
	private Short verticalAlignment;
	private Color backgroundColor;
	private Font font;
	private String format;

	private Style() {}

	private Style copy() {
		Style result = new Style();
		result.alignment = alignment;
		result.verticalAlignment = verticalAlignment;
		result.backgroundColor = backgroundColor;
		result.font = font;
		result.format = format;
		return result;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		Style style = (Style) o;
		return Objects.equals(alignment, style.alignment) &&
				Objects.equals(verticalAlignment, style.verticalAlignment) &&
				backgroundColor == style.backgroundColor &&
				Objects.equals(font, style.font) &&
				Objects.equals(format, style.format);
	}

	@Override
	public int hashCode() {
		return Objects.hash(alignment, verticalAlignment, backgroundColor, font, format);
	}

	public Style alignLeft() {
		return alignment(CellStyle.ALIGN_LEFT);
	}

	public Style alignCenter() {
		return alignment(CellStyle.ALIGN_CENTER);
	}

	public Style alignRight() {
		return alignment(CellStyle.ALIGN_RIGHT);
	}

	public Style alignTop() {
		return verticalAlignment(CellStyle.VERTICAL_TOP);
	}

	public Style alignMiddle() {
		return verticalAlignment(CellStyle.VERTICAL_CENTER);
	}

	public Style alignBottom() {
		return verticalAlignment(CellStyle.VERTICAL_BOTTOM);
	}

	public Style backgroundColor(Color backgroundColor) {
		Style result = copy();
		result.backgroundColor = backgroundColor;
		return result;
	}

	public Style font(Font font) {
		Style result = copy();
		result.font = font;
		return result;
	}

	public Style format(String format) {
		Style result = copy();
		result.format = format;
		return result;
	}

	private Style alignment(short alignment) {
		Style result = copy();
		result.alignment = alignment;
		return result;
	}

	private Style verticalAlignment(short verticalAlignment) {
		Style result = copy();
		result.verticalAlignment = verticalAlignment;
		return result;
	}

	public Short getAlignment() {
		return alignment;
	}

	public Short getVerticalAlignment() {
		return verticalAlignment;
	}

	public Color getBackgroundColor() {
		return backgroundColor;
	}

	public Font getFont() {
		return font;
	}

	public String getFormat() {
		return format;
	}
}
