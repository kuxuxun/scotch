package com.github.kuxuxun.scotch.excel.cell;

import java.math.BigDecimal;
import java.util.Date;


import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScPos;

public class ScCell {

	private final ScSheet sheet;
	private final ScPos pos;
	private final Cell poiCell;

	private final ScStyle style;

	public String getTextValue() {
		return this.poiCell.getRichStringCellValue().getString();
	}

	public BigDecimal getNumericValue() {
		return new BigDecimal(poiCell.getNumericCellValue());
	}

	private ScCell(ScPos pos, ScSheet wb) {
		this.sheet = wb;
		this.pos = pos;

		this.poiCell = getCellFromSheet(pos, wb);
		this.style = getStyleFromSheet(poiCell, pos, wb);
	}

	private Cell getCellFromSheet(ScPos pos, ScSheet sheet) {
		Cell cell = null;
		Row row = sheet.getPoiRowAt(pos.getRow());
		if ((cell = row.getCell(pos.getCol())) == null) {
			cell = row.createCell(pos.getCol());
		}
		return cell;
	}

	private ScStyle getStyleFromSheet(Cell poiCell, ScPos pos, ScSheet sheet) {

		CellStyle style = poiCell.getCellStyle();
		ScStyle s = ScStyle.createNew(sheet.getRowHeightOf(pos.getRow()), sheet
				.getColWidthOf(pos.getCol()));

		s.setCellStyleWithoutFont(ScWithoutFontStyle.createWithPoiStyle(style));

		s.setFontStyle(sheet.getWorkBook().createFont(style));

		return s;
	}

	public static ScCell getAt(ScPos pos, ScSheet sheet) {
		ScCell c = sheet.getWorkBook().getCellFromCache(pos);
		if (c != null) {
			return c;
		}

		return new ScCell(pos, sheet);
	}

	public Date getDateValue() {
		return poiCell.getDateCellValue();
	}

	public CellType getCellType() {
		return CellType.getTypeFromNumberAs(poiCell.getCellType());
	}

	public String getFormula() {
		return poiCell.getCellFormula();
	}

	public void lock(boolean b) {
		poiCell.getCellStyle().setLocked(b);
	}

	public ScStyle getStyle() {
		return style;
	}

	public void setTextWithStyle(String value, ScStyle cellStyle) {
		setStyle(cellStyle);
		poiCell.setCellValue(new HSSFRichTextString(value));

	}

	public void setNumberToCell(BigDecimal value, ScStyle cellStyle) {

		setStyle(cellStyle);
		poiCell.setCellValue(value.doubleValue());
	}

	public void setFunctionWithStyle(String formula, ScStyle cellStyle) {

		setStyle(cellStyle);
		poiCell.setCellFormula(formula);
	}

	public enum CellType {

		BLANK(Cell.CELL_TYPE_BLANK), BOOLEAN(Cell.CELL_TYPE_BOOLEAN), FORMULA(
				Cell.CELL_TYPE_FORMULA), STRING(Cell.CELL_TYPE_STRING), NUMERIC(
				Cell.CELL_TYPE_NUMERIC), ERROR(Cell.CELL_TYPE_ERROR), UNKNOWN(
				-99);
		private final int typeNumber;

		private CellType(int typeNumber) {
			this.typeNumber = typeNumber;
		}

		public static CellType getTypeFromNumberAs(int number) {

			for (CellType each : CellType.values()) {
				if (each.typeNumber == number) {
					return each;
				}
			}
			return UNKNOWN;
		}

		public boolean is(CellType t) {
			return this == t;
		}

		public boolean isBlank() {
			return this == BLANK;
		}

		public boolean isUnknown() {
			return this == UNKNOWN;
		}
	}

	public ScPos getPos() {
		return pos;
	}

	public Cell getPoiCell() {
		return poiCell;
	}

	public void setStyle(ScStyle style) {

		setStyleWithoutHeightAndWidthTo(style);

		sheet.getPoiRowAt(pos.getRow()).setHeight(
				style.getRowHeight().shortValue());

		sheet.getPoiSheet().setColumnWidth(poiCell.getColumnIndex(),
				style.getColWidth().shortValue());
	}

	public void setStyleWithoutHeightAndWidthTo(ScStyle style) {

		ScWithoutFontStyle targetCellStyle = sheet.getWorkBook()
				.getStyleFromCacheOrCreateNew(style).getStyleWithoutFont();

		targetCellStyle.setStyleTo(poiCell);

	}

	public boolean isDateFormatted() {
		return DateUtil.isCellDateFormatted(this.poiCell);
	}

}
