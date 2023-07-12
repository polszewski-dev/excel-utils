package polszewski.excel.writer.style;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * User: POlszewski
 * Date: 2023-07-13
 */
class StyleTest {
	@Test
	void alignLeft() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.alignLeft();

		assertThat(changed.alignment()).isEqualTo(HorizontalAlignment.LEFT);
		assertThat(original.alignment()).isNull();
	}

	@Test
	void alignCenter() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.alignCenter();

		assertThat(changed.alignment()).isEqualTo(HorizontalAlignment.CENTER);
		assertThat(original.alignment()).isNull();
	}

	@Test
	void alignRight() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.alignRight();

		assertThat(changed.alignment()).isEqualTo(HorizontalAlignment.RIGHT);
		assertThat(original.alignment()).isNull();
	}

	@Test
	void alignTop() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.alignTop();

		assertThat(changed.verticalAlignment()).isEqualTo(VerticalAlignment.TOP);
		assertThat(original.verticalAlignment()).isNull();
	}

	@Test
	void alignMiddle() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.alignMiddle();

		assertThat(changed.verticalAlignment()).isEqualTo(VerticalAlignment.CENTER);
		assertThat(original.verticalAlignment()).isNull();
	}

	@Test
	void alignBottom() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.alignBottom();

		assertThat(changed.verticalAlignment()).isEqualTo(VerticalAlignment.BOTTOM);
		assertThat(original.verticalAlignment()).isNull();
	}

	@Test
	void backgroundColor() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.backgroundColor(Color.AQUA);

		assertThat(changed.backgroundColor()).isEqualTo(Color.AQUA);
		assertThat(original.backgroundColor()).isNull();
	}

	@Test
	void font() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.font(Font.DEFAULT);

		assertThat(changed.font()).isEqualTo(Font.DEFAULT);
		assertThat(original.font()).isNull();
	}

	@Test
	void format() {
		Style original = new Style(null, null, null, null, null);
		Style changed = original.format("0.00");

		assertThat(changed.format()).isEqualTo("0.00");
		assertThat(original.format()).isNull();
	}
}