package com.github.kuxuxun.scotch.excel;

import static org.hamcrest.core.Is.is;
import static org.junit.Assert.assertThat;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Before;
import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.area.ScRange;
import com.github.kuxuxun.scotch.excel.cell.ScStyle;
import com.github.kuxuxun.scotch.excel.cell.ScWithoutFontStyle;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.BordersInSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders.Border;

public class BordersInSheetTest extends TestHelper {

	private BordersInSheet b;

	@Before
	public void setup() {
		b = new BordersInSheet();
	}

	@Test
	public void testBorderAroundACell() {

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.BOTTOM,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.TOP, Border
				.thinAutoColor());

		assertThat(b.isAroundedIn(new ScPos("B3"), Border.thinAutoColor()),
				is(true));

	}

	@Test
	public void testBorderAroundACellFault() {

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.BOTTOM,
				Border.thinAutoColor());

		assertThat(b.isAroundedIn(new ScPos("B3"), Border.thinAutoColor()),
				is(false));

	}

	@Test
	public void testBorderAroundSomeCell() {

		setBorderAroundInB3toC4(b);

		assertThat(b.isAroundedIn(
				new ScRange(new ScPos("B3"), new ScPos("C4")), Border
						.thinAutoColor()), is(true));

	}

	private void setBorderAroundInB3toC4(BordersInSheet b) {
		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.TOP, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("B4"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("B4"), BordersInSheet.Destination.BOTTOM,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("C3"), BordersInSheet.Destination.TOP, Border
				.thinAutoColor());
		b.setBorderdAt(new ScPos("C3"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("C4"), BordersInSheet.Destination.BOTTOM,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("C4"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());
	}

	@Test
	public void testBorderAroundLazy() {

		setBorderAroundInLazyInB3ToC4(b);

		assertThat(b.isAroundedIn(
				new ScRange(new ScPos("B3"), new ScPos("C4")), Border
						.thinAutoColor()), is(true));

	}

	@Test
	public void testEqualLookAndFeel() {

		setBorderAroundInLazyInB3ToC4(b);

		BordersInSheet counter = new BordersInSheet();
		setBorderAroundInB3toC4(counter);

		assertThat(b.hasEqualLookAndFeelWith(counter), is(true));

	}

	@Test
	public void testEqualLookAndFeelFault() {

		setBorderAroundInLazyInB3ToC4(b);

		BordersInSheet counter = new BordersInSheet();
		setBorderAroundInB3toC4(counter);

		counter.setBorderdAt(new ScPos("Z4"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());

		assertThat(b.hasEqualLookAndFeelWith(counter), is(false));

	}

	private void setBorderAroundInLazyInB3ToC4(BordersInSheet b) {
		b.setBorderdAt(new ScPos("B2"), BordersInSheet.Destination.BOTTOM,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("C2"), BordersInSheet.Destination.BOTTOM,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("B5"), BordersInSheet.Destination.TOP, Border
				.thinAutoColor());
		b.setBorderdAt(new ScPos("C5"), BordersInSheet.Destination.TOP, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("A3"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());
		b.setBorderdAt(new ScPos("A4"), BordersInSheet.Destination.RIGHT,
				Border.thinAutoColor());

		b.setBorderdAt(new ScPos("D3"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());

		b.setBorderdAt(new ScPos("D4"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());
	}

	@Test
	public void testRegisterCell() {
		b.setBorderdAt(new ScPos("B3"), BordersInSheet.Destination.LEFT, Border
				.thinAutoColor());

		assertThat(b.hasBorderedAt(new ScPos("B3"),
				BordersInSheet.Destination.LEFT, Border.thinAutoColor()),
				is(true));

	}

	@Test
	public void testRegister() throws IllegalAccessException {
		ScStyle style = createMockCellStyle();

		b.getFromCellStyle(new ScPos("B3"), style);

		assertThat(b.hasBorderedAt(new ScPos("B3"),
				BordersInSheet.Destination.LEFT, Border.thinBlack()), is(true));

		assertThat(b.hasBorderedAt(new ScPos("B3"),
				BordersInSheet.Destination.RIGHT, Border.thinBlack()), is(true));

		assertThat(b.hasBorderedAt(new ScPos("B3"),
				BordersInSheet.Destination.TOP, Border.thinBlack()), is(true));

		assertThat(b.hasBorderedAt(new ScPos("B3"),
				BordersInSheet.Destination.BOTTOM, Border.thinBlack()),
				is(true));
	}

	private ScStyle createMockCellStyle() throws IllegalAccessException {
		ScWithoutFontStyle styleWithoutFont = ScWithoutFontStyle
				.createWithPoiStyle(new MockCellStyle());

		styleWithoutFont.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setLeftBorderColor(new HSSFColor.BLACK().getIndex());

		styleWithoutFont.setBorderRight(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setRightBorderColor(new HSSFColor.BLACK().getIndex());

		styleWithoutFont.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setBottomBorderColor(new HSSFColor.BLACK().getIndex());

		styleWithoutFont.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setTopBorderColor(new HSSFColor.BLACK().getIndex());

		ScStyle style = ScStyle.createNew(decimal(0), decimal(0));
		setValueToFieldOf(style, "cellStyleWithoutFont", styleWithoutFont);
		return style;
	}

	public class MockCellStyle implements CellStyle {

		public void cloneStyleFrom(CellStyle source) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public short getAlignment() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderBottom() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderLeft() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderRight() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderTop() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBottomBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getDataFormat() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public String getDataFormatString() {
			// TODO 自動生成されたメソッド・スタブ
			return null;
		}

		public short getFillBackgroundColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public Color getFillBackgroundColorColor() {
			// TODO 自動生成されたメソッド・スタブ
			return null;
		}

		public short getFillForegroundColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public Color getFillForegroundColorColor() {
			// TODO 自動生成されたメソッド・スタブ
			return null;
		}

		public short getFillPattern() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getFontIndex() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public boolean getHidden() {
			// TODO 自動生成されたメソッド・スタブ
			return false;
		}

		public short getIndention() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getIndex() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getLeftBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public boolean getLocked() {
			// TODO 自動生成されたメソッド・スタブ
			return false;
		}

		public short getRightBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getRotation() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getTopBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getVerticalAlignment() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public boolean getWrapText() {
			// TODO 自動生成されたメソッド・スタブ
			return false;
		}

		public void setAlignment(short align) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderBottom(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderLeft(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderRight(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderTop(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBottomBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setDataFormat(short fmt) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFillBackgroundColor(short bg) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFillForegroundColor(short bg) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFillPattern(short fp) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFont(Font font) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setHidden(boolean hidden) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setIndention(short indent) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setLeftBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setLocked(boolean locked) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setRightBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setRotation(short rotation) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setTopBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setVerticalAlignment(short align) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setWrapText(boolean wrapped) {
			// TODO 自動生成されたメソッド・スタブ

		}

	}

}
