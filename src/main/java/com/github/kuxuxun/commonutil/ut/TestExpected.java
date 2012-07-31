package com.github.kuxuxun.commonutil.ut;

import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.ScWorkbook;

public class TestExpected {

	private final ScWorkbook wb;

	private TestExpected(String fileName) throws IOException {
		wb = new ScWorkbook(new HSSFWorkbook(this.getClass().getClassLoader()
				.getResourceAsStream("expected/" + fileName)));
	}

	public static ScSheet getFirstSheetOf(String fileName) {
		TestExpected e;
		try {
			e = new TestExpected(fileName);
		} catch (Exception ex) {
			throw new IllegalArgumentException("シートの取得に失敗しました。 " + fileName);
		}

		return e.wb.getFirstSheet();
	}

	public static ScWorkbook fileOf(String fileName) throws IOException {
		return new TestExpected(fileName).wb;
	}

}
