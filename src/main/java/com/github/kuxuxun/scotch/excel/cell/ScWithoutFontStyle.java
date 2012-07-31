package com.github.kuxuxun.scotch.excel.cell;

import java.util.Arrays;

import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class ScWithoutFontStyle implements Cloneable {

	private final CellStyle poiCellStyle;

	private short alignment;
	private short fillBackgroundColor;
	private short fillForegroundColor;
	private short fillPattern;
	private boolean hidden;
	private short indention;
	private boolean locked;
	private short borderBottom;
	private short borderLeft;
	private short borderRight;
	private short borderTop;
	private short bottomBorderColor;
	private short rightBorderColor;
	private short leftBorderColor;
	private short rotation;
	private short topBorderColor;
	private short verticalAlignment;
	private boolean wrapText;

	private ScWithoutFontStyle(CellStyle poiCellStyle) {
		this.poiCellStyle = poiCellStyle;

		copyStylePropetiesFrom(poiCellStyle);
	}

	public static ScWithoutFontStyle createWithPoiStyle(CellStyle poiCellStyle) {
		return new ScWithoutFontStyle(poiCellStyle);
	}

	public String getDataFormatString() {
		return this.poiCellStyle.getDataFormatString();
	}

	public String key() {
		int[] key = styleElements();
		return Arrays.toString(key) + poiCellStyle.getDataFormatString();
	}

	private int[] styleElements() {
		int[] key = new int[19];

		key[0] = getAlignment();
		key[1] = getBorderBottom();
		key[2] = getBorderLeft();
		key[3] = getBorderRight();
		key[4] = getBorderTop();
		key[5] = getBottomBorderColor();

		key[7] = getFillBackgroundColor();
		key[8] = getFillForegroundColor();
		key[9] = getFillPattern();
		if (isHidden()) {
			key[10] = 1;
		}
		key[11] = getIndention();
		key[12] = getLeftBorderColor();

		if (isLocked()) {
			key[13] = 1;
		}

		key[14] = getRightBorderColor();
		key[15] = getRotation();
		key[16] = getTopBorderColor();
		key[17] = getVerticalAlignment();
		if (poiCellStyle.getWrapText()) {
			key[18] = 1;
		}
		return key;
	}

	/**
	 * 
	 * セルの書式情報を対象にコピーします。
	 * <p>
	 * 呼び出し側で{@link Workbook#createCellStyle()}により作成されたcellStyleをこのメソッドに渡すことにより、
	 * 作成されたcellStyleに書式がコピーされます。
	 * 
	 * @param dst
	 */
	protected void copyCellStyleTo(ScWithoutFontStyle dst) {

		dst.poiCellStyle.setAlignment(getAlignment());

		dst.poiCellStyle.setFillBackgroundColor(getFillBackgroundColor());
		dst.poiCellStyle.setFillForegroundColor(getFillForegroundColor());
		dst.poiCellStyle.setFillPattern(getFillPattern());

		dst.poiCellStyle.setHidden(isHidden());
		dst.poiCellStyle.setIndention(getIndention());
		dst.poiCellStyle.setLocked(isLocked());
		dst.poiCellStyle.setRotation(getRotation());
		dst.poiCellStyle.setVerticalAlignment(getVerticalAlignment());
		dst.poiCellStyle.setWrapText(isWrapText());

		dst.poiCellStyle.setDataFormat(poiCellStyle.getDataFormat());

		dst.poiCellStyle.setBorderBottom(getBorderBottom());
		dst.poiCellStyle.setBorderLeft(getBorderLeft());
		dst.poiCellStyle.setBorderRight(getBorderRight());
		dst.poiCellStyle.setBorderTop(getBorderTop());
		dst.poiCellStyle.setBottomBorderColor(getBottomBorderColor());
		dst.poiCellStyle.setLeftBorderColor(getLeftBorderColor());
		dst.poiCellStyle.setRightBorderColor(getRightBorderColor());
		dst.poiCellStyle.setTopBorderColor(getTopBorderColor());

	}

	@Override
	public int hashCode() {

		HashCodeBuilder builder = new HashCodeBuilder();
		return builder.append(styleElements()).append(
				poiCellStyle.getDataFormatString()).toHashCode();
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj) {
			return true;
		}
		if (!(obj instanceof ScWithoutFontStyle)) {
			return false;
		}

		ScWithoutFontStyle c = (ScWithoutFontStyle) obj;

		EqualsBuilder builder = new EqualsBuilder();
		builder.append(this.styleElements(), c.styleElements());
		builder.append(this.poiCellStyle.getDataFormatString(), c.poiCellStyle
				.getDataFormatString());

		return builder.isEquals();

	}

	public ScColor getBgColor() {
		return ScColor.getFromColorIndex(fillBackgroundColor);
	}

	public ScColor getFgColor() {
		return ScColor.getFromColorIndex(fillForegroundColor);
	}

	private void copyStylePropetiesFrom(CellStyle poiCellStyle) {
		alignment = poiCellStyle.getAlignment();
		fillBackgroundColor = poiCellStyle.getFillBackgroundColor();
		fillForegroundColor = poiCellStyle.getFillForegroundColor();
		fillPattern = poiCellStyle.getFillPattern();
		hidden = poiCellStyle.getHidden();
		indention = poiCellStyle.getIndention();
		locked = poiCellStyle.getLocked();
		borderBottom = poiCellStyle.getBorderBottom();
		borderLeft = poiCellStyle.getBorderLeft();
		borderRight = poiCellStyle.getBorderRight();
		borderTop = poiCellStyle.getBorderTop();
		bottomBorderColor = poiCellStyle.getBottomBorderColor();
		rightBorderColor = poiCellStyle.getRightBorderColor();
		leftBorderColor = poiCellStyle.getLeftBorderColor();
		rotation = poiCellStyle.getRotation();
		topBorderColor = poiCellStyle.getTopBorderColor();
		verticalAlignment = poiCellStyle.getVerticalAlignment();
		wrapText = poiCellStyle.getWrapText();

	}

	private void copyStylePropetiesTo(CellStyle style) {
		style.setAlignment(alignment);
		style.setFillBackgroundColor(fillBackgroundColor);
		style.setFillForegroundColor(fillForegroundColor);
		style.setFillPattern(fillPattern);
		style.setHidden(hidden);
		style.setIndention(indention);
		style.setLocked(locked);
		style.setBorderBottom(borderBottom);
		style.setBorderLeft(borderLeft);
		style.setBorderRight(borderRight);
		style.setBorderTop(borderTop);
		style.setBottomBorderColor(bottomBorderColor);
		style.setRightBorderColor(rightBorderColor);
		style.setLeftBorderColor(leftBorderColor);
		style.setRotation(rotation);
		style.setTopBorderColor(topBorderColor);
		style.setVerticalAlignment(verticalAlignment);
		style.setWrapText(wrapText);

	}

	public short getAlignment() {
		return alignment;
	}

	public void setAlignment(short alignment) {
		this.alignment = alignment;
	}

	public short getFillBackgroundColor() {
		return fillBackgroundColor;
	}

	public void setFillBackgroundColor(short fillBackgroundColor) {
		this.fillBackgroundColor = fillBackgroundColor;
	}

	public short getFillForegroundColor() {
		return fillForegroundColor;
	}

	public void setFillForegroundColor(short fillForegroundColor) {
		this.fillForegroundColor = fillForegroundColor;
	}

	public short getFillPattern() {
		return fillPattern;
	}

	public void setFillPattern(short fillPattern) {
		this.fillPattern = fillPattern;
	}

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	public short getIndention() {
		return indention;
	}

	public void setIndention(short indention) {
		this.indention = indention;
	}

	public boolean isLocked() {
		return locked;
	}

	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	public short getBorderBottom() {
		return borderBottom;
	}

	public void setBorderBottom(short borderBottom) {
		this.borderBottom = borderBottom;
	}

	public short getBorderLeft() {
		return borderLeft;
	}

	public void setBorderLeft(short borderLeft) {
		this.borderLeft = borderLeft;
	}

	public short getBorderRight() {
		return borderRight;
	}

	public void setBorderRight(short borderRight) {
		this.borderRight = borderRight;
	}

	public short getBorderTop() {
		return borderTop;
	}

	public void setBorderTop(short borderTop) {
		this.borderTop = borderTop;
	}

	public short getBottomBorderColor() {
		return bottomBorderColor;
	}

	public void setBottomBorderColor(short bottomBorderColor) {
		this.bottomBorderColor = bottomBorderColor;
	}

	public short getRightBorderColor() {
		return rightBorderColor;
	}

	public void setRightBorderColor(short rightBorderColor) {
		this.rightBorderColor = rightBorderColor;
	}

	public short getLeftBorderColor() {
		return leftBorderColor;
	}

	public void setLeftBorderColor(short leftBorderColor) {
		this.leftBorderColor = leftBorderColor;
	}

	public short getRotation() {
		return rotation;
	}

	public void setRotation(short rotation) {
		this.rotation = rotation;
	}

	public short getTopBorderColor() {
		return topBorderColor;
	}

	public void setTopBorderColor(short topBorderColor) {
		this.topBorderColor = topBorderColor;
	}

	public short getVerticalAlignment() {
		return verticalAlignment;
	}

	public void setVerticalAlignment(short verticalAlignment) {
		this.verticalAlignment = verticalAlignment;
	}

	public boolean isWrapText() {
		return wrapText;
	}

	public void setWrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	public void setStyleTo(Cell cell) {

		this.copyStylePropetiesTo(poiCellStyle);
		cell.setCellStyle(this.poiCellStyle);
	}

	@Override
	public String toString() {

		StringBuilder b = new StringBuilder();

		b.append(blacket("alignment :" + alignment));
		b.append(blacket("fillBackgroundColor :" + fillBackgroundColor));
		b.append(blacket("fillForegroundColor :" + fillForegroundColor));
		b.append(blacket("fillPattern :" + fillPattern));
		b.append(blacket("hidden :" + hidden));
		b.append(blacket("indention :" + indention));
		b.append(blacket("locked :" + locked));
		b.append(blacket("borderBottom :" + borderBottom));
		b.append(blacket("borderLeft :" + borderLeft));
		b.append(blacket("borderRight :" + borderRight));
		b.append(blacket("borderTop :" + borderTop));
		b.append(blacket("bottomBorderColor :" + bottomBorderColor));
		b.append(blacket("rightBorderColor :" + rightBorderColor));
		b.append(blacket("leftBorderColor :" + leftBorderColor));
		b.append(blacket("rotation :" + rotation));
		b.append(blacket("topBorderColor :" + topBorderColor));
		b.append(blacket("verticalAlignment :" + verticalAlignment));
		b.append(blacket("wrapText :" + wrapText));
		return b.toString();

	}

	private static String blacket(String str) {
		return "[" + str + "]";
	}
}
