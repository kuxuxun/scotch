package com.github.kuxuxun.scotch.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.kuxuxun.scotch.excel.area.ScArea;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScFont;
import com.github.kuxuxun.scotch.excel.cell.ScStyle;
import com.github.kuxuxun.scotch.excel.cell.ScWithoutFontStyle;
import com.github.kuxuxun.scotch.excel.util.ReflectionHelper;

public class ScWorkbook extends ReflectionHelper {

	private final Workbook poiWorkbook;
	private final StyleCache styleCache = new StyleCache();

	public ScWorkbook(Workbook workbook) {
		this.poiWorkbook = workbook;
	}

	private final Map<ScPos, ScCell> cellCache = new HashMap<ScPos, ScCell>();

        public ScCell putCellToCache(ScCell c) {
		this.cellCache.put(c.getPos(), c);
		return c;
	}

	public ScCell getCellFromCache(ScPos pos) {
		// TODO sheetに移す
		return this.cellCache.get(pos);
	}

	public Workbook getPoiWorkBook() {
		// TODO protected に
		return poiWorkbook;
	}

	// FIXME UT作成後、このメソッドをリファクタリング
	public ScStyle getStyleFromCacheOrCreateNew(ScStyle concernedStyle) {

		ScStyle cached = styleCache.get(concernedStyle);

		if (cached != null) {
			return cached;
		}

		CellStyle poiCellStyle = getPoiWorkBook().createCellStyle();

		ScStyle result = ScStyle.createNew(concernedStyle.getRowHeight(),
				concernedStyle.getColWidth());

		if (concernedStyle.getStyleWithoutFont() != null) {

			ScWithoutFontStyle newCellStyleWithoutFont = ScWithoutFontStyle
					.createWithPoiStyle(poiCellStyle);

			concernedStyle.copyCellStyleTo(newCellStyleWithoutFont);

			DataFormat dataFormat = poiWorkbook.createDataFormat();
			poiCellStyle.setDataFormat(dataFormat.getFormat(concernedStyle
					.getStyleWithoutFont().getDataFormatString()));

		}

		if (concernedStyle.getFontStyle() != null) {
			Font targetCellFont = concernedStyle.getFontStyle().findFontFromWb(
					this);

			if (targetCellFont == null) {
				targetCellFont = getPoiWorkBook().createFont();
				concernedStyle.copyFontStyleTo(new ScFont(targetCellFont));
			}
			poiCellStyle.setFont(targetCellFont);
			result.setFontStyle(concernedStyle.getFontStyle());
		}

		result.setCellStyleWithoutFont(ScWithoutFontStyle
				.createWithPoiStyle(poiCellStyle));

		styleCache.put(result);

		return result;
	}

	public class StyleCache {
		private final Map<String, ScStyle> cache = new HashMap<String, ScStyle>();

		public void put(ScStyle cellStyle) {
			cache.put(cellStyle.keyOfStyle(), cellStyle);
		}

		public ScStyle get(ScStyle target) {
			return this.cache.get(target.keyOfStyle());
		}
	}

	public void setRepeatPrintRow(int startRow, int endRow) {
		this.poiWorkbook
				.setRepeatingRowsAndColumns(0, -1, -1, startRow, endRow);
	}

	/**
	 * @return
	 * @deprecated 未実装です、、
	 */
	@Deprecated
	public ScArea getRepeatingRow() {

		throw new UnsupportedOperationException("未実装");

		// 実装用のメモ
		// Name name = this.poiWorkbook.getNameAt(this.poiWorkbook
		// .getNameIndex("Print_Titles"));

		// Looks like POI does not have a way to get
		// to the repeating rows and columns. So you
		// probably will need to modify source.
		//
		// Here's some information that may help you:

		// 1. User code:
		//
		// fileInputStream = // get the input
		// HSSFWorkbook w = new HSSFWorkbook(fileInputStream);
		// HSSFName name = w.getNameAt(w.getNameIndex("Print_Titles"));
		// // Note: "Print_Titles" is a fixed value that you will
		// // need to use to get to the NameRecord for Repeating
		// // Rows and Cols
		//
		// 2. And then to get the repeating rows and columns:
		//
		// AreaReference repeatingRows
		// = new AreaReference(name.getUnionReference()[0]);
		// AreaReference repeatingCols
		// = new AreaReference(name.getUnionReference()[1]);
		//
		// 3. Ofcourse, HSSFName.getUnionReference() does not exist, so you will
		// have to add it as:
		//
		// public String[] getUnionReference() {
		// return name.getAreaUnionReference(book);
		// }
		//
		// 4. And since NameRecord.getAreaUnionReference(book) does not exist
		// either, you can add it in NameRecord.java as:
		//
		// public String[] getAreaUnionReference(Workbook book) {
		// String[] result = null;
		// if (field_13_name_definition != null
		// && !field_13_name_definition.isEmpty()) {
		// Stack ptgStack = (Stack) field_13_name_definition.clone();
		//
		// switch (ptgStack.size()) {
		// case 4:
		// Ptg ptg = (Ptg) ptgStack.pop();
		// if (ptg.getClass() == UnionPtg.class) {
		// result = new String[2];
		// Area3DPtg aptg0 = (Area3DPtg) ptgStack.pop();
		// Area3DPtg aptg1 = (Area3DPtg) ptgStack.pop();
		// result[0] = aptg0.toFormulaString(book);
		// result[1] = aptg1.toFormulaString(book);
		// }
		// break;
		// case 1:
		// result = new String[1];
		// // TODO: check class
		// Area3DPtg aptg0 = (Area3DPtg) ptgStack.pop();
		// result[0] = aptg0.toFormulaString(book);
		// break;
		// }
		// }
		// return result;
		// }
		//
		// Hope that helps,
		// ~ amol
		//
		// PS: At the moment I'm not sure thats the best way to
		// do it so I wont be putting up a patch for it, but
		// give it a try and see how it works for you. I have tried
		// it for basic cases and it seems to work both for reading
		// from POI generated workbooks and excel generated workbooks.
	}

	public Font findFont(short boldWeight, short color, short fontHeight,
			String name, boolean italic, boolean strikeout, short typeOffset,
			byte underline) {

		return this.poiWorkbook.findFont(boldWeight, color, fontHeight, name,
				italic, strikeout, typeOffset, underline);

	}

	public ScFont createFont(CellStyle style) {
		return new ScFont(this.poiWorkbook.getFontAt(style.getFontIndex()));
	}

	public ScSheet getFirstSheet() {
		return this.getSheetAt(0);
	}

	public ScSheet getSheetAt(int sheetNo) {
		return ScSheet.getAt(sheetNo, this);
	}

	public ScSheet createSheet(String sheetName) {
		return ScSheet.createNew(sheetName, this);
	}

	public List<ScSheet> getAllSheets() {
		List<ScSheet> ss = new ArrayList<ScSheet>();

		int num = poiWorkbook.getNumberOfSheets();
		for (int i = 0; i < num; i++) {
			ss.add(ScSheet.getAt(i, this));
		}

		return ss;

	}

}
