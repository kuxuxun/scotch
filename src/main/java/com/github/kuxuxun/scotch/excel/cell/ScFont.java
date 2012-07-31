package com.github.kuxuxun.scotch.excel.cell;

import java.awt.FontMetrics;
import java.awt.Label;


import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.kuxuxun.scotch.excel.ScWorkbook;

public class ScFont {

	private final Font poiFont;

	@Deprecated
	public int getFontWidthOfNumber() {
		// このメソッドの動作の正当性が確認されていません。
		FontMetrics metrics = new Label().getFontMetrics(toAwtFont());
		return metrics.charWidth('0');

	}

	public java.awt.Font toAwtFont() {

		int style = java.awt.Font.PLAIN;
		if (isBold()) {
			style |= java.awt.Font.BOLD;
		}
		if (isItalic()) {
			style |= java.awt.Font.ITALIC;
		}

		return new java.awt.Font(getFontName(), style, getFontSize());
	}

	public ScFont(Font poiFont) {
		this.poiFont = poiFont;
	}

	public ScColor getFontColor() {
		return ScColor.getFromColorIndex(poiFont.getColor());
	}

	public String getFontName() {
		return poiFont.getFontName();
	}

	public int getFontSize() {
		return poiFont.getFontHeightInPoints();
	}

	public boolean isBold() {
		return (poiFont.getBoldweight() == Font.BOLDWEIGHT_BOLD);
	}

	public boolean isItalic() {
		return poiFont.getItalic();
	}

	public boolean isUnderlinedWithSingle() {
		return (poiFont.getUnderline() == Font.U_SINGLE);
	}

	public boolean isUnderlinedWithDouble() {
		return (poiFont.getUnderline() == Font.U_DOUBLE);
	}

	/**
	 * ワークブックから作成済みのfontを検索し返却します、該当のフォントが作成されていない場合、nullを返却します
	 * 
	 * @param searching
	 * @return
	 */
	public Font findFontFromWb(ScWorkbook wb) {

		return wb.findFont(poiFont.getBoldweight(), poiFont.getColor(), poiFont
				.getFontHeight(), poiFont.getFontName(), poiFont.getItalic(),
				poiFont.getStrikeout(), poiFont.getTypeOffset(), poiFont
						.getUnderline());
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
	protected void copyFontStyleTo(ScFont dst) {

		dst.poiFont.setBoldweight(poiFont.getBoldweight());
		dst.poiFont.setCharSet(poiFont.getCharSet());
		dst.poiFont.setColor(poiFont.getColor());
		dst.poiFont.setFontHeight(poiFont.getFontHeight());
		dst.poiFont.setFontName(poiFont.getFontName());
		dst.poiFont.setItalic(poiFont.getItalic());
		dst.poiFont.setStrikeout(poiFont.getStrikeout());
		dst.poiFont.setTypeOffset(poiFont.getTypeOffset());
		dst.poiFont.setUnderline(poiFont.getUnderline());
	}

	public String key() {
		StringBuilder fontKey = new StringBuilder();

		fontKey.append(Short.toString(poiFont.getBoldweight()));
		fontKey.append(Short.toString(poiFont.getColor()));
		fontKey.append(Short.toString(poiFont.getFontHeight()));
		fontKey.append(poiFont.getFontName());
		if (poiFont.getItalic()) {
			fontKey.append("1");
		} else {
			fontKey.append("0");
		}

		if (poiFont.getStrikeout()) {
			fontKey.append("1");
		} else {
			fontKey.append("0");
		}

		fontKey.append(Short.toString(poiFont.getTypeOffset()));
		fontKey.append(Short.toString(poiFont.getUnderline()));

		return fontKey.toString();

	}

	@Override
	public int hashCode() {

		HashCodeBuilder builder = new HashCodeBuilder();
		return builder.append(key()).toHashCode();
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj) {
			return true;
		}
		if (!(obj instanceof ScFont)) {
			return false;
		}

		ScFont c = (ScFont) obj;

		EqualsBuilder builder = new EqualsBuilder();
		builder.append(this.key(), c.key());

		return builder.isEquals();

	}

	public Font getPoiFont() {
		return poiFont;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();

		builder.append("[boldweight:" + poiFont.getBoldweight() + "]");
		builder.append("[font height:" + poiFont.getFontHeight() + "]");
		builder.append("[color:" + poiFont.getColor() + "]");
		builder.append("[italic :" + poiFont.getItalic() + "]");
		builder.append("[strike out:" + poiFont.getStrikeout() + "]");
		builder.append("[type offset :"
				+ Short.toString(poiFont.getTypeOffset()) + "]");
		builder.append("[underline:" + Short.toString(poiFont.getUnderline())
				+ "]");

		builder.append("[font:name" + poiFont.getFontName() + "]");

		return builder.toString();
	}

}
