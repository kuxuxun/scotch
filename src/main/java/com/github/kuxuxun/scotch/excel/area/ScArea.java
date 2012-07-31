package com.github.kuxuxun.scotch.excel.area;


import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScStyle;

public abstract class ScArea {

	public abstract int getTopRow();

	public abstract int getBottomRow();

	public abstract int getLeftCol();

	public abstract int getRightCol();

	public abstract String toReference();

	public void fillWith(ScStyle style, ScSheet sheet) {
		for (int row = getTopRow(); row <= getBottomRow(); row++) {
			for (int col = getLeftCol(); col <= getRightCol(); col++) {

				ScCell c = sheet.getCellAt(row, col);
				c.setStyleWithoutHeightAndWidthTo(style);
			}
		}

	}

	public void setColoredBorderAround(ScSheet sheet, IndexedColors color) {
		setBorderAround(sheet, CellStyle.BORDER_THIN, color);
	}

	public void setBorderAround(ScSheet sheet) {
		setBorderAround(sheet, CellStyle.BORDER_THIN);
	}

	public void setBorderAround(ScSheet sheet, final short borderStyle) {
		setBorderAround(sheet, borderStyle, IndexedColors.BLACK);
	}

	public void setBorderAround(ScSheet sheet, final short borderStyle,
			IndexedColors color) {

		setBorderLeft(sheet, borderStyle, color);
		setBorderTop(sheet, borderStyle, color);
		setBorderRight(sheet, borderStyle, color);
		setBorderBottom(sheet, borderStyle, color);
	}

	public void setBorderLeft(ScSheet sheet, final short borderStyle,
			IndexedColors color) {

		final int col = getLeftCol();
		for (int row = getTopRow(); row <= getBottomRow(); row++) {

			ScCell c = sheet.getCellAt(row, col);

			ScStyle targetStyle = c.getStyle();

			targetStyle.getStyleWithoutFont().setBorderLeft(borderStyle);
			targetStyle.getStyleWithoutFont().setLeftBorderColor(
					color.getIndex());

			c.setStyle(targetStyle);

		}
	}

	public void setBorderRight(ScSheet sheet, final short borderStyle,
			IndexedColors color) {

		final int col = getRightCol();
		for (int row = getTopRow(); row <= getBottomRow(); row++) {

			ScCell c = sheet.getCellAt(row, col);

			ScStyle targetStyle = c.getStyle();
			targetStyle.getStyleWithoutFont().setBorderRight(borderStyle);
			targetStyle.getStyleWithoutFont().setRightBorderColor(
					color.getIndex());

			c.setStyle(targetStyle);

		}
	}

	public void setBorderTop(ScSheet sheet, final short borderStyle,
			IndexedColors color) {

		final int row = getTopRow();
		for (int col = getLeftCol(); col <= getRightCol(); col++) {
			ScCell c = sheet.getCellAt(row, col);

			ScStyle targetStyle = c.getStyle();
			targetStyle.getStyleWithoutFont().setBorderTop(borderStyle);
			targetStyle.getStyleWithoutFont().setTopBorderColor(
					color.getIndex());

			c.setStyle(targetStyle);
		}
	}

	public void setBorderBottom(ScSheet sheet, final short borderStyle,
			IndexedColors color) {

		final int row = getBottomRow();
		for (int col = getLeftCol(); col <= getRightCol(); col++) {

			ScCell c = sheet.getCellAt(row, col);

			ScStyle targetStyle = c.getStyle();

			targetStyle.getStyleWithoutFont().setBorderBottom(borderStyle);
			targetStyle.getStyleWithoutFont().setBottomBorderColor(
					color.getIndex());

			c.setStyle(targetStyle);

		}

	}

}
