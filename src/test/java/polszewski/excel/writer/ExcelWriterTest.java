package polszewski.excel.writer;

import org.junit.jupiter.api.Test;
import polszewski.excel.reader.templates.ExcelParser;
import polszewski.excel.reader.templates.ExcelSheetParser;
import polszewski.excel.writer.style.Color;
import polszewski.excel.writer.style.Font;
import polszewski.excel.writer.style.Style;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * User: POlszewski
 * Date: 2023-07-13
 */
class ExcelWriterTest {
	@Test
	void writeAndRead() throws IOException {
		Path tempFile = Files.createTempFile(null, null);

		createXlsFile(tempFile);

		TestExcelParser parser = new TestExcelParser(new TestExcelSheetParser() {
			@Override
			protected void readSingleRow() {
				if (getCurrentRowIdx() == 1) {
					assertThat(stringCol.getString()).isEqualTo("1st-row");
					assertThat(doubleCol.getDouble()).isEqualTo(12.5);
				} else if (getCurrentRowIdx() == 2) {
					assertThat(stringCol.getString()).isEqualTo("2nd-row");
					assertThat(doubleCol.getDouble()).isEqualTo(25);
				} else if (getCurrentRowIdx() == 3) {
					assertThat(stringCol.getString()).isEmpty();
					assertThat(doubleCol.getDouble()).isEqualTo(37.5);
				}
				++rowsRead;
			}
		}, tempFile);

		parser.readFromXls();

		assertThat(rowsRead).isEqualTo(3);
	}

	static int rowsRead;

	static void createXlsFile(Path tempFile) throws IOException {
		Font headerFont = Font.DEFAULT
				.name("Arial")
				.size(12)
				.bold(true)
				.italic(true)
				.color(Color.WHITE);

		Style headerStyle = Style.DEFAULT
				.font(headerFont)
				.backgroundColor(Color.BLACK)
				.alignCenter()
				.alignMiddle();

		Style doubleStyle = Style.DEFAULT
				.font(Font.DEFAULT)
				.format("0.00");

		ExcelWriter writer = new ExcelWriter();

		writer.open("test");

		writer.setCell("string", headerStyle);
		writer.setCell("double", headerStyle);
		writer.nextRow();

		writer.setCell("1st-row");
		writer.setCell(12.5, doubleStyle);
		writer.nextRow();

		writer.setCell("2nd-row");
		writer.setCell(25, doubleStyle);
		writer.nextRow();

		writer.setCell("");
		writer.setCell(37.5);

		writer.save(tempFile.toString());
	}

	static class TestExcelParser extends ExcelParser {
		final TestExcelSheetParser sheetParser;
		final Path path;

		TestExcelParser(TestExcelSheetParser sheetParser, Path path) {
			this.sheetParser = sheetParser;
			this.path = path;
		}

		@Override
		protected InputStream getExcelInputStream() throws FileNotFoundException {
			return new FileInputStream(path.toFile());
		}

		@Override
		protected Stream<ExcelSheetParser> getSheetParsers() {
			return Stream.of(sheetParser);
		}
	}

	static abstract class TestExcelSheetParser extends ExcelSheetParser {
		ExcelColumn stringCol = column("string");
		ExcelColumn doubleCol = column("double");

		TestExcelSheetParser() {
			super("test");
		}

		@Override
		protected ExcelColumn getColumnIndicatingOptionalRow() {
			return stringCol;
		}
	}
}