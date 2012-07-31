package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertEquals;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;

import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.style.ScMargin;
import com.github.kuxuxun.scotch.excel.style.ScPaperSize;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.PageFitting;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.PageNumberPrint;

public class AssertPageSetting extends TestHelper {

	private final ScSheet sheet;

	private AssertPageSetting(ScSheet sheet) {
		this.sheet = sheet;
	}

	public static AssertPageSetting that(ScSheet sheet) {
		return new AssertPageSetting(sheet);
	}

	private ScMargin marginOf(ScSheet s) {
		return ScMargin.getFromSheet(s);
	}

	public void headerMarginIs(BigDecimal expected) throws Exception {

		assertCompareEquals("ヘッダマージン:", expected, marginOf(sheet).getHeader());

	}

	public void footerMarginIs(BigDecimal expected) throws Exception {

		assertCompareEquals("フッタマージン:", expected, marginOf(sheet).getFooter());
	}

	public void leftMarginIs(BigDecimal expected) throws Exception {

		assertCompareEquals("左マージン:", expected, marginOf(sheet).getLeft());
	}

	public void rightMarginIs(BigDecimal expected) throws Exception {

		assertCompareEquals("右マージン:", expected, marginOf(sheet).getRight());
	}

	public void topMarginIs(BigDecimal expected) throws Exception {

		assertCompareEquals("上マージン:", expected, marginOf(sheet).getTop());
	}

	public void bottomMarginIs(BigDecimal expected) throws Exception {

		assertCompareEquals("下マージン:", expected, marginOf(sheet).getBottom());
	}

	public void pagePaperSizeIs(String sizeSortName) throws Exception {
	    assertEquals("用紙サイズ:", ScPaperSize.getFromShortName(sizeSortName),
			 sheet.getPaperSize());
	}

	public void pageOrientation(String orientation) throws Exception {
	    assertEquals("用紙の向き:", orientation, sheet.getPaperOrientation());
	}

	public void printPageFitTo(PageFitting expected) throws Exception {

	    PageFitting f = PageFitting.getFromSheet(sheet);
	    
	    assertEquals("ページに合わせて印刷:", expected, f);
	}

	public void printPageNumberSettingIs(PageNumberPrint expected)
			throws FileNotFoundException, IOException {

		assertEquals("ページ番号出力:", expected, PageNumberPrint.getFromSheet(sheet));

	}
	// FIXME ヘッダ繰り返しタイトルのアサート実装

}
