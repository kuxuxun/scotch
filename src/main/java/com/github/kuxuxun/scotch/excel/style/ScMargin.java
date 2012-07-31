package com.github.kuxuxun.scotch.excel.style;

import java.math.BigDecimal;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;

import com.github.kuxuxun.scotch.excel.ScSheet;

public class ScMargin {

	private BigDecimal top;
	private BigDecimal bottom;
	private BigDecimal left;
	private BigDecimal right;

	private BigDecimal header;
	private BigDecimal footer;

	public static BigDecimal inchToCentiMeter(BigDecimal inch) {

		// FIXME 変換方法よく分からん
		BigDecimal d = inchToMillMeters(inch).divide(BigDecimal.TEN);

		if (d.setScale(3, BigDecimal.ROUND_HALF_UP).compareTo(
				new BigDecimal("1.905")) == 0) {
			return new BigDecimal(2);
		}

		return d.setScale(1, BigDecimal.ROUND_HALF_UP);
	}

	private static BigDecimal inchToMillMeters(BigDecimal inInch) {

		return inInch.multiply(new BigDecimal(25.4));

	}

	public static ScMargin getFromSheet(ScSheet sheet) {

		ScMargin m = new ScMargin();
		PrintSetup printSetting = sheet.getPoiSheet().getPrintSetup();

		m.header = inchToCentiMeter(new BigDecimal(printSetting
				.getHeaderMargin()));
		m.footer = inchToCentiMeter(new BigDecimal(printSetting
				.getFooterMargin()));

		Sheet poisheet = sheet.getPoiSheet();
		m.top = inchToCentiMeter(new BigDecimal(poisheet
				.getMargin(HSSFSheet.TopMargin)));
		m.bottom = inchToCentiMeter(new BigDecimal(poisheet
				.getMargin(HSSFSheet.BottomMargin)));
		m.left = inchToCentiMeter(new BigDecimal(poisheet
				.getMargin(HSSFSheet.LeftMargin)));
		m.right = inchToCentiMeter(new BigDecimal(poisheet
				.getMargin(HSSFSheet.RightMargin)));

		return m;
	}

	public BigDecimal getTop() {
		return top;
	}

	public void setTop(BigDecimal top) {
		this.top = top;
	}

	public BigDecimal getBottom() {
		return bottom;
	}

	public void setBottom(BigDecimal bottom) {
		this.bottom = bottom;
	}

	public BigDecimal getLeft() {
		return left;
	}

	public void setLeft(BigDecimal left) {
		this.left = left;
	}

	public BigDecimal getRight() {
		return right;
	}

	public void setRight(BigDecimal right) {
		this.right = right;
	}

	public BigDecimal getHeader() {
		return header;
	}

	public void setHeader(BigDecimal header) {
		this.header = header;
	}

	public BigDecimal getFooter() {
		return footer;
	}

	public void setFooter(BigDecimal footer) {
		this.footer = footer;
	}

}
