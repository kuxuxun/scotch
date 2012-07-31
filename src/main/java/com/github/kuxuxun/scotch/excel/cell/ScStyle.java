package com.github.kuxuxun.scotch.excel.cell;

import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Workbook;

public class ScStyle implements Cloneable {

	private ScWithoutFontStyle cellStyleWithoutFont;
	private ScFont fontStyle;

	private final BigDecimal rowHeight;
	private final BigDecimal colWidth;

	private ScStyle(BigDecimal rowHeight, BigDecimal colWidth) {
		this.rowHeight = rowHeight;
		this.colWidth = colWidth;
	}

	public static ScStyle createNew(BigDecimal rowHeight, BigDecimal colWidth) {
		return new ScStyle(rowHeight, colWidth);
	}

	public ScWithoutFontStyle getStyleWithoutFont() {
		return cellStyleWithoutFont;
	}

	public void setCellStyleWithoutFont(ScWithoutFontStyle cellBehavior) {
		this.cellStyleWithoutFont = cellBehavior;
	}

	public ScFont getFontStyle() {
		return fontStyle;
	}

	public void setFontStyle(ScFont fontStyle) {
		this.fontStyle = fontStyle;
	}

	/**
	 * セルの書式情報を対象にコピーします。
	 * <p>
	 * 呼び出し側で{@link Workbook#createCellStyle()}により作成されたcellStyleをこのメソッドに渡すことにより、
	 * 作成されたcellStyleに書式がコピーされます。
	 * 
	 * @param src
	 * @param dst
	 */
	public void copyCellStyleTo(ScWithoutFontStyle dst) {

		dst.copyCellStyleTo(dst);
		cellStyleWithoutFont.copyCellStyleTo(dst);
	}

	/**
	 * フォント情報を対象にコピーします。
	 * 
	 * <p>
	 * 呼び出し側で{@link Workbook#createFont()}により作成されたfontをこのメソッドに渡すことにより、
	 * 作成されたfontに書式がコピーされます。
	 * 
	 * @param src
	 * @param dst
	 */
	public void copyFontStyleTo(ScFont dst) {
		fontStyle.copyFontStyleTo(dst);
	}

	public BigDecimal getRowHeight() {
		return rowHeight;
	}

	public BigDecimal getColWidth() {
		return colWidth;
	}

	public String keyOfStyle() {

		String fontKey = "";

		if (fontStyle != null) {
			fontKey = this.fontStyle.key();
		}

		return cellStyleWithoutFont.key() + fontKey;
	}
}
