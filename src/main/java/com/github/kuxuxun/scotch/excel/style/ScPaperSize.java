package com.github.kuxuxun.scotch.excel.style;

import org.apache.poi.ss.usermodel.PrintSetup;

public enum ScPaperSize {

	// TODO org.apache.poi.ss.usermodel.PaperSizeとの整合性の取り方を考える
	LETTER("LETTER", PrintSetup.LETTER_PAPERSIZE), LETTER_SMALL("LETTER_S",
			PrintSetup.LETTER_SMALL_PAGESIZE), TABLOID("TABLOID",
			PrintSetup.TABLOID_PAPERSIZE), LEDGER("LEDGER",
			PrintSetup.LEDGER_PAPERSIZE), LEGAL("LEGAL",
			PrintSetup.LEGAL_PAPERSIZE), STATEMENT("STATEMENT",
			PrintSetup.STATEMENT_PAPERSIZE), EXECUTIVE("EXECUTIVE",
			PrintSetup.EXECUTIVE_PAPERSIZE), A3("A3", PrintSetup.A3_PAPERSIZE), A4(
			"A4", PrintSetup.A4_PAPERSIZE), A4_SMALL("A4_S",
			PrintSetup.A4_SMALL_PAPERSIZE), A5("A5", PrintSetup.A5_PAPERSIZE), B4(
			"B4", PrintSetup.B4_PAPERSIZE), B5("B5", PrintSetup.B5_PAPERSIZE), FOLIO(
			"FOLIO", PrintSetup.FOLIO8_PAPERSIZE), QUARTO("QUARTO",
			PrintSetup.QUARTO_PAPERSIZE);

	private final String shortName;
	private final int indexOfHssfPaperSize;

	private ScPaperSize(String shortName, int indexOfHssfPaperSize) {

		this.shortName = shortName;
		this.indexOfHssfPaperSize = indexOfHssfPaperSize;

	}

	public static ScPaperSize getFromShortName(String shortName) {
		for (ScPaperSize each : values()) {
			if (each.getShortName().equals(shortName)) {
				return each;
			}
		}
		throw new IllegalArgumentException("不正な用紙サイズ名です : " + shortName);
	}

	public static ScPaperSize getFromPoiPaperIndex(short ps) {
		for (ScPaperSize each : values()) {
			if (each.indexOfHssfPaperSize == ps) {
				return each;
			}
		}
		throw new IllegalArgumentException("未対応な用紙サイズです : " + ps);
	}

	public String getShortName() {
		return shortName;
	}

}
