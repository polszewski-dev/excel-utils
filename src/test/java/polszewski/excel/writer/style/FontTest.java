package polszewski.excel.writer.style;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * User: POlszewski
 * Date: 2023-07-13
 */
class FontTest {
	@Test
	void name() {
		Font original = new Font(null, null, null, null, null);
		Font changed = original.name("Arial");

		assertThat(changed.name()).isEqualTo("Arial");
		assertThat(original.name()).isNull();
	}

	@Test
	void size() {
		Font original = new Font(null, null, null, null, null);
		Font changed = original.size(12);

		assertThat(changed.size()).isEqualTo(12);
		assertThat(original.size()).isNull();
	}

	@Test
	void color() {
		Font original = new Font(null, null, null, null, null);
		Font changed = original.color(Color.AQUA);

		assertThat(changed.color()).isEqualTo(Color.AQUA);
		assertThat(original.color()).isNull();
	}

	@Test
	void bold() {
		Font original = new Font(null, null, null, null, null);
		Font changed = original.bold(true);

		assertThat(changed.bold()).isTrue();
		assertThat(original.bold()).isNull();
	}

	@Test
	void italic() {
		Font original = new Font(null, null, null, null, null);
		Font changed = original.italic(true);

		assertThat(changed.italic()).isTrue();
		assertThat(original.italic()).isNull();
	}

	@Test
	void normalFont() {
		Font original = Font.DEFAULT.bold(true).italic(true);

		assertThat(original.bold()).isTrue();
		assertThat(original.italic()).isTrue();

		Font changed = original.normalFont();

		assertThat(changed.bold()).isFalse();
		assertThat(changed.italic()).isFalse();
	}
}