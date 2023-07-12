package polszewski.excel.reader.templates;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Set;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static polszewski.excel.reader.templates.ExcelParserTest.TestEnum.*;

/**
 * User: POlszewski
 * Date: 2023-07-13
 */
class ExcelParserTest {
	@Test
	void testRequiredPrimitiveValues() throws IOException {
		readIt(new TestExcelSheetParser(1) {
			@Override
			void readOnlyOneRow() {
				String stringVal = stringCol.getString();
				int integerVal = integerCol.getInteger();
				Integer nullableIntegerVal = integerCol.getNullableInteger();
				double doubleVal = doubleCol.getDouble();
				Double nullableDoubleVal = doubleCol.getNullableDouble();
				boolean booleanVal = booleanCol.getBoolean();
				TestEnum enumVal = enumCol.getEnum(TestEnum::valueOf);

				assertThat(stringVal).isEqualTo("test");
				assertThat(integerVal).isEqualTo(1);
				assertThat(nullableIntegerVal).isEqualTo(1);
				assertThat(doubleVal).isEqualTo(1.25);
				assertThat(nullableDoubleVal).isEqualTo(1.25);
				assertThat(booleanVal).isTrue();
				assertThat(enumVal).isEqualTo(THREE);
			}
		});
	}

	@Test
	void testRequiredPrimitiveValuesAsLists() throws IOException {
		readIt(new TestExcelSheetParser(1) {
			@Override
			void readOnlyOneRow() {
				List<String> stringListVal = stringCol.getList(x -> x);
				List<Integer> integerListVal = integerCol.getList(Integer::valueOf);
				List<Double> doubleListVal = doubleCol.getList(Double::valueOf);
				List<Boolean> booleanListVal = booleanCol.getList(Boolean::valueOf);
				List<TestEnum> enumListVal = enumCol.getList(TestEnum::valueOf);

				assertThat(stringListVal).isEqualTo(List.of("test"));
				assertThat(integerListVal).isEqualTo(List.of(1));
				assertThat(doubleListVal).isEqualTo(List.of(1.25));
				assertThat(booleanListVal).isEqualTo(List.of(true));
				assertThat(enumListVal).isEqualTo(List.of(THREE));
			}
		});
	}

	@Test
	void testRequiredListValues() throws IOException {
		readIt(new TestExcelSheetParser(1) {
			@Override
			void readOnlyOneRow() {
				List<String> stringListVal = stringListCol.getList(x -> x);
				List<Integer> integerListVal = integerListCol.getList(Integer::valueOf);
				List<Double> doubleListVal = doubleListCol.getList(Double::valueOf);
				List<Boolean> booleanListVal = booleanListCol.getList(Boolean::valueOf);
				List<TestEnum> enumListVal = enumListCol.getList(TestEnum::valueOf);

				assertThat(stringListVal).isEqualTo(List.of("a", "b", "c"));
				assertThat(integerListVal).isEqualTo(List.of(1, 2, 3));
				assertThat(doubleListVal).isEqualTo(List.of(0.5, 1.5, 2.5));
				assertThat(booleanListVal).isEqualTo(List.of(true, false));
				assertThat(enumListVal).isEqualTo(List.of(THREE, TWO, ONE));
			}
		});
	}

	@Test
	void testRequiredSetValues() throws IOException {
		readIt(new TestExcelSheetParser(1) {
			@Override
			void readOnlyOneRow() {
				Set<String> stringListVal = stringListCol.getSet(x -> x);
				Set<Integer> integerListVal = integerListCol.getSet(Integer::valueOf);
				Set<Double> doubleListVal = doubleListCol.getSet(Double::valueOf);
				Set<Boolean> booleanListVal = booleanListCol.getSet(Boolean::valueOf);
				Set<TestEnum> enumListVal = enumListCol.getSet(TestEnum::valueOf);

				assertThat(stringListVal).isEqualTo(Set.of("a", "b", "c"));
				assertThat(integerListVal).isEqualTo(Set.of(1, 2, 3));
				assertThat(doubleListVal).isEqualTo(Set.of(0.5, 1.5, 2.5));
				assertThat(booleanListVal).isEqualTo(Set.of(true, false));
				assertThat(enumListVal).isEqualTo(Set.of(THREE, TWO, ONE));
			}
		});
	}

	@Test
	void testOptionalPrimitiveValues() throws IOException {
		readIt(new TestExcelSheetParser(2) {
			@Override
			void readOnlyOneRow() {
				assertThatThrownBy(() -> stringCol.getString()).isInstanceOf(IllegalArgumentException.class);
				assertThatThrownBy(() -> integerCol.getInteger()).isInstanceOf(IllegalArgumentException.class);
				assertThatThrownBy(() -> doubleCol.getDouble()).isInstanceOf(IllegalArgumentException.class);
				assertThatThrownBy(() -> enumCol.getEnum(TestEnum::valueOf)).isInstanceOf(IllegalArgumentException.class);

				assertThat(booleanCol.getBoolean()).isFalse();

				String stringVal = stringCol.getString("empty");
				int integerVal = integerCol.getInteger(-1);
				Integer nullableIntegerVal = integerCol.getNullableInteger();
				double doubleVal = doubleCol.getDouble(-2.0);
				Double nullableDoubleVal = doubleCol.getNullableDouble();
				TestEnum enumVal = enumCol.getEnum(TestEnum::valueOf, ONE);

				assertThat(stringVal).isEqualTo("empty");
				assertThat(integerVal).isEqualTo(-1);
				assertThat(nullableIntegerVal).isNull();
				assertThat(doubleVal).isEqualTo(-2);
				assertThat(nullableDoubleVal).isNull();
				assertThat(enumVal).isEqualTo(ONE);
			}
		});
	}

	@Test
	void testOptionalListValues() throws IOException {
		readIt(new TestExcelSheetParser(2) {
			@Override
			void readOnlyOneRow() {
				List<String> stringListVal = stringListCol.getList(x -> x);
				List<Integer> integerListVal = integerListCol.getList(Integer::valueOf);
				List<Double> doubleListVal = doubleListCol.getList(Double::valueOf);
				List<Boolean> booleanListVal = booleanListCol.getList(Boolean::valueOf);
				List<TestEnum> enumListVal = enumListCol.getList(TestEnum::valueOf);

				assertThat(stringListVal).isEqualTo(List.of());
				assertThat(integerListVal).isEqualTo(List.of());
				assertThat(doubleListVal).isEqualTo(List.of());
				assertThat(booleanListVal).isEqualTo(List.of());
				assertThat(enumListVal).isEqualTo(List.of());
			}
		});
	}

	@Test
	void testSpecialValues() throws IOException {
		readIt(new TestExcelSheetParser(3) {
			@Override
			void readOnlyOneRow() {
				assertThat(formulaCol.getString()).isEqualTo("1+2");
				assertThat(integerCol.getBoolean()).isTrue();
				assertThat(booleanCol.getBoolean()).isFalse();
			}
		});
	}

	@Test
	void testOptionalColumns() throws IOException {
		readIt(new TestExcelSheetParser(3) {
			@Override
			void readOnlyOneRow() {
				assertThatThrownBy(() -> notExistingCol.getString()).isInstanceOf(IllegalArgumentException.class);
				assertThatThrownBy(() -> notExistingCol.getString("empty")).isInstanceOf(IllegalArgumentException.class);

				assertThatThrownBy(() -> optionalCol.getString()).isInstanceOf(IllegalArgumentException.class);
				assertThat(optionalCol.getString("empty")).isEqualTo("empty");
			}
		});
	}

	@Test
	void testColumnDefinition() {
		TestExcelSheetParser parser = new TestExcelSheetParser(1) {
			@Override
			void readOnlyOneRow() {

			}
		};

		assertThat(parser.stringCol.getName()).isEqualTo("string");
		assertThat(parser.stringCol.isOptional()).isFalse();

		assertThat(parser.optionalCol.getName()).isEqualTo("optional");
		assertThat(parser.optionalCol.isOptional()).isTrue();
	}

	@Test
	void test() throws IOException {
		readIt(new TestExcelSheetParser(1) {
			@Override
			void readOnlyOneRow() {
				var columnNames = getColumnNames(".*_list");

				assertThat(columnNames).isEqualTo(List.of("string_list", "integer_list", "double_list", "boolean_list", "enum_list"));
			}
		});
	}

	static void readIt(TestExcelSheetParser sheetParser) throws IOException {
		new TestExcelParser(sheetParser).readFromXls();

		assertThat(sheetParser.rowRead).isTrue();
	}

	static class TestExcelParser extends ExcelParser {
		final TestExcelSheetParser sheetParser;

		TestExcelParser(TestExcelSheetParser sheetParser) {
			this.sheetParser = sheetParser;
		}

		@Override
		protected InputStream getExcelInputStream() {
			return fromResourcePath("/parser_test.xls");
		}

		@Override
		protected Stream<ExcelSheetParser> getSheetParsers() {
			return Stream.of(sheetParser);
		}
	}

	static abstract class TestExcelSheetParser extends ExcelSheetParser {
		ExcelColumn stringCol = column("string");
		ExcelColumn integerCol = column("integer");
		ExcelColumn doubleCol = column("double");
		ExcelColumn booleanCol = column("boolean");
		ExcelColumn enumCol = column("enum");

		ExcelColumn stringListCol = column("string_list");
		ExcelColumn integerListCol = column("integer_list");
		ExcelColumn doubleListCol = column("double_list");
		ExcelColumn booleanListCol = column("boolean_list");
		ExcelColumn enumListCol = column("enum_list");

		ExcelColumn formulaCol = column("formula");

		ExcelColumn notExistingCol = column("not-existing");
		ExcelColumn optionalCol = column("optional", true);

		final int onlyRowIndex;
		boolean rowRead;

		TestExcelSheetParser(int onlyRowIndex) {
			super("data");
			this.onlyRowIndex = onlyRowIndex;
		}

		@Override
		protected ExcelColumn getColumnIndicatingOptionalRow() {
			return null;
		}

		@Override
		protected void readSingleRow() {
			if (getCurrentRowIdx() == onlyRowIndex) {
				readOnlyOneRow();
				this.rowRead = true;
			}
		}

		abstract void readOnlyOneRow();
	}

	enum TestEnum {
		ONE, TWO, THREE
	}
}
