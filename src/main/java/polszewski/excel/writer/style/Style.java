package polszewski.excel.writer.style;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public record Style(
		HorizontalAlignment alignment,
		VerticalAlignment verticalAlignment,
		Color backgroundColor,
		Font font,
		String format
) {
	public static final Style DEFAULT = new Style(null, null, null, null, null);

	public Style alignLeft() {
		return alignment(HorizontalAlignment.LEFT);
	}

	public Style alignCenter() {
		return alignment(HorizontalAlignment.CENTER);
	}

	public Style alignRight() {
		return alignment(HorizontalAlignment.RIGHT);
	}

	public Style alignTop() {
		return verticalAlignment(VerticalAlignment.TOP);
	}

	public Style alignMiddle() {
		return verticalAlignment(VerticalAlignment.CENTER);
	}

	public Style alignBottom() {
		return verticalAlignment(VerticalAlignment.BOTTOM);
	}

	public Style backgroundColor(Color backgroundColor) {
		return new Style(alignment, verticalAlignment, backgroundColor, font, format);
	}

	public Style font(Font font) {
		return new Style(alignment, verticalAlignment, backgroundColor, font, format);
	}

	public Style format(String format) {
		return new Style(alignment, verticalAlignment, backgroundColor, font, format);
	}

	private Style alignment(HorizontalAlignment alignment) {
		return new Style(alignment, verticalAlignment, backgroundColor, font, format);
	}

	private Style verticalAlignment(VerticalAlignment verticalAlignment) {
		return new Style(alignment, verticalAlignment, backgroundColor, font, format);
	}
}
