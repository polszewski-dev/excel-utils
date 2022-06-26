package polszewski.excel.writer.style;

import org.apache.poi.hssf.util.HSSFColor;

/**
 * User: POlszewski
 * Date: 2019-01-10
 */
public enum Color {
	AUTOMATIC(new HSSFColor.AUTOMATIC()),
	LIGHT_CORNFLOWER_BLUE(new HSSFColor.LIGHT_CORNFLOWER_BLUE()),
	ROYAL_BLUE(new HSSFColor.ROYAL_BLUE()),
	CORAL(new HSSFColor.CORAL()),
	ORCHID(new HSSFColor.ORCHID()),
	MAROON(new HSSFColor.MAROON()),
	LEMON_CHIFFON(new HSSFColor.LEMON_CHIFFON()),
	CORNFLOWER_BLUE(new HSSFColor.CORNFLOWER_BLUE()),
	WHITE(new HSSFColor.WHITE()),
	LAVENDER(new HSSFColor.LAVENDER()),
	PALE_BLUE(new HSSFColor.PALE_BLUE()),
	LIGHT_TURQUOISE(new HSSFColor.LIGHT_TURQUOISE()),
	LIGHT_GREEN(new HSSFColor.LIGHT_GREEN()),
	LIGHT_YELLOW(new HSSFColor.LIGHT_YELLOW()),
	TAN(new HSSFColor.TAN()),
	ROSE(new HSSFColor.ROSE()),
	GREY_25_PERCENT(new HSSFColor.GREY_25_PERCENT()),
	PLUM(new HSSFColor.PLUM()),
	SKY_BLUE(new HSSFColor.SKY_BLUE()),
	TURQUOISE(new HSSFColor.TURQUOISE()),
	BRIGHT_GREEN(new HSSFColor.BRIGHT_GREEN()),
	YELLOW(new HSSFColor.YELLOW()),
	GOLD(new HSSFColor.GOLD()),
	PINK(new HSSFColor.PINK()),
	GREY_40_PERCENT(new HSSFColor.GREY_40_PERCENT()),
	VIOLET(new HSSFColor.VIOLET()),
	LIGHT_BLUE(new HSSFColor.LIGHT_BLUE()),
	AQUA(new HSSFColor.AQUA()),
	SEA_GREEN(new HSSFColor.SEA_GREEN()),
	LIME(new HSSFColor.LIME()),
	LIGHT_ORANGE(new HSSFColor.LIGHT_ORANGE()),
	RED(new HSSFColor.RED()),
	GREY_50_PERCENT(new HSSFColor.GREY_50_PERCENT()),
	BLUE_GREY(new HSSFColor.BLUE_GREY()),
	BLUE(new HSSFColor.BLUE()),
	TEAL(new HSSFColor.TEAL()),
	GREEN(new HSSFColor.GREEN()),
	DARK_YELLOW(new HSSFColor.DARK_YELLOW()),
	ORANGE(new HSSFColor.ORANGE()),
	DARK_RED(new HSSFColor.DARK_RED()),
	GREY_80_PERCENT(new HSSFColor.GREY_80_PERCENT()),
	INDIGO(new HSSFColor.INDIGO()),
	DARK_BLUE(new HSSFColor.DARK_BLUE()),
	DARK_TEAL(new HSSFColor.DARK_TEAL()),
	DARK_GREEN(new HSSFColor.DARK_GREEN()),
	OLIVE_GREEN(new HSSFColor.OLIVE_GREEN()),
	BROWN(new HSSFColor.BROWN()),
	BLACK(new HSSFColor.BLACK());

	private final HSSFColor hssfColor;

	Color(HSSFColor hssfColor) {
		this.hssfColor = hssfColor;
	}

	public short getIndex() {
		return hssfColor.getIndex();
	}
}
