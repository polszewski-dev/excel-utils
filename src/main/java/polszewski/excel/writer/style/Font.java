package polszewski.excel.writer.style;

import java.util.Objects;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public class Font {
	public static final Font DEFAULT = new Font();

	private String name;
	private Integer size;
	private Color color;
	private Boolean bold;
	private Boolean italic;

	private Font() {}

	private Font copy() {
		Font result = new Font();
		result.name = name;
		result.size = size;
		result.color = color;
		result.bold = bold;
		result.italic = italic;
		return result;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (!(o instanceof Font)) return false;
		Font font = (Font) o;
		return Objects.equals(name, font.name) && Objects.equals(size, font.size) && color == font.color && Objects.equals(bold, font.bold) && Objects.equals(italic, font.italic);
	}

	@Override
	public int hashCode() {
		return Objects.hash(name, size, color, bold, italic);
	}

	public Font name(String name) {
		Font result = copy();
		result.name = name;
		return result;
	}

	public Font size(Integer size) {
		Font result = copy();
		result.size = size;
		return result;
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

	public String getName() {
		return name;
	}

	public Integer getSize() {
		return size;
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
