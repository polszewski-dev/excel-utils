package polszewski.excel.writer.style;

import java.util.Objects;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public class Font {
	public static final Font DEFAULT = new Font();

	private Color color;
	private Boolean bold;
	private Boolean italic;

	private Font() {}

	private Font copy() {
		Font result = new Font();
		result.color = color;
		result.bold = bold;
		return result;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		Font font = (Font) o;
		return color == font.color &&
				Objects.equals(bold, font.bold);
	}

	@Override
	public int hashCode() {
		return Objects.hash(color, bold);
	}

	public Font color(Color color) {
		Font result = copy();
		result.color = color;
		return result;
	}

	public Font bold() {
		Font result = copy();
		result.bold = true;
		return result;
	}

	public Font italic() {
		Font result = copy();
		result.italic = true;
		return result;
	}

	public Font normalFont() {
		Font result = copy();
		result.bold = false;
		result.italic = false;
		return result;
	}

	public Color getColor() {
		return color;
	}

	public Boolean getBold() {
		return bold;
	}

	public Boolean getItalic() {
		return italic;
	}
}
