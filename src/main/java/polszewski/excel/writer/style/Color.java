package polszewski.excel.writer.style;

import static org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public enum Color {
	AUTOMATIC(HSSFColorPredefined.AUTOMATIC),
	LIGHT_CORNFLOWER_BLUE(HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE),
	ROYAL_BLUE(HSSFColorPredefined.ROYAL_BLUE),
	CORAL(HSSFColorPredefined.CORAL),
	ORCHID(HSSFColorPredefined.ORCHID),
	MAROON(HSSFColorPredefined.MAROON),
	LEMON_CHIFFON(HSSFColorPredefined.LEMON_CHIFFON),
	CORNFLOWER_BLUE(HSSFColorPredefined.CORNFLOWER_BLUE),
	WHITE(HSSFColorPredefined.WHITE),
	LAVENDER(HSSFColorPredefined.LAVENDER),
	PALE_BLUE(HSSFColorPredefined.PALE_BLUE),
	LIGHT_TURQUOISE(HSSFColorPredefined.LIGHT_TURQUOISE),
	LIGHT_GREEN(HSSFColorPredefined.LIGHT_GREEN),
	LIGHT_YELLOW(HSSFColorPredefined.LIGHT_YELLOW),
	TAN(HSSFColorPredefined.TAN),
	ROSE(HSSFColorPredefined.ROSE),
	GREY_25_PERCENT(HSSFColorPredefined.GREY_25_PERCENT),
	PLUM(HSSFColorPredefined.PLUM),
	SKY_BLUE(HSSFColorPredefined.SKY_BLUE),
	TURQUOISE(HSSFColorPredefined.TURQUOISE),
	BRIGHT_GREEN(HSSFColorPredefined.BRIGHT_GREEN),
	YELLOW(HSSFColorPredefined.YELLOW),
	GOLD(HSSFColorPredefined.GOLD),
	PINK(HSSFColorPredefined.PINK),
	GREY_40_PERCENT(HSSFColorPredefined.GREY_40_PERCENT),
	VIOLET(HSSFColorPredefined.VIOLET),
	LIGHT_BLUE(HSSFColorPredefined.LIGHT_BLUE),
	AQUA(HSSFColorPredefined.AQUA),
	SEA_GREEN(HSSFColorPredefined.SEA_GREEN),
	LIME(HSSFColorPredefined.LIME),
	LIGHT_ORANGE(HSSFColorPredefined.LIGHT_ORANGE),
	RED(HSSFColorPredefined.RED),
	GREY_50_PERCENT(HSSFColorPredefined.GREY_50_PERCENT),
	BLUE_GREY(HSSFColorPredefined.BLUE_GREY),
	BLUE(HSSFColorPredefined.BLUE),
	TEAL(HSSFColorPredefined.TEAL),
	GREEN(HSSFColorPredefined.GREEN),
	DARK_YELLOW(HSSFColorPredefined.DARK_YELLOW),
	ORANGE(HSSFColorPredefined.ORANGE),
	DARK_RED(HSSFColorPredefined.DARK_RED),
	GREY_80_PERCENT(HSSFColorPredefined.GREY_80_PERCENT),
	INDIGO(HSSFColorPredefined.INDIGO),
	DARK_BLUE(HSSFColorPredefined.DARK_BLUE),
	DARK_TEAL(HSSFColorPredefined.DARK_TEAL),
	DARK_GREEN(HSSFColorPredefined.DARK_GREEN),
	OLIVE_GREEN(HSSFColorPredefined.OLIVE_GREEN),
	BROWN(HSSFColorPredefined.BROWN),
	BLACK(HSSFColorPredefined.BLACK);

	private final HSSFColorPredefined hssfColor;

	Color(HSSFColorPredefined hssfColor) {
		this.hssfColor = hssfColor;
	}

	public short getIndex() {
		return hssfColor.getIndex();
	}
}
