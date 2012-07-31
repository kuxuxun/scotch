package com.github.kuxuxun.scotch.excel;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;


import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.area.ScRange;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.style.ScPaperSize;
import com.github.kuxuxun.scotch.excel.util.ReflectionHelper;

public class ScSheet extends ReflectionHelper {

	private final Sheet poiSheet;
	private final ScWorkbook wb;
	private int currentRow = 0;

	private ScSheet(Sheet s, ScWorkbook wb) {
		this.poiSheet = s;
		this.wb = wb;
	}

	public static ScSheet getAt(int sheetNo, ScWorkbook wb) {
		return new ScSheet(wb.getPoiWorkBook().getSheetAt(sheetNo), wb);
	}

	public static ScSheet createNew(String sheetName, ScWorkbook wb) {
		ScSheet s = new ScSheet(wb.getPoiWorkBook().createSheet(sheetName), wb);
		s.poiSheet.setDisplayGridlines(false);
		return s;
	}

	public Row getPoiRowAt(int rowPos) {
		Row row;
		if ((row = this.poiSheet.getRow(rowPos)) == null) {
			row = this.poiSheet.createRow(rowPos);
		}

		return row;
	}

	public void setColWidthOf(int colPos, int height) {
		poiSheet.setColumnWidth(colPos, height);
	}

	public void setRowHeightOf(int rowPos, short height) {
		Row row = this.getPoiRowAt(rowPos);
		row.setHeight(height);
	}

	public BigDecimal getRowHeightOf(int row) {
		return new BigDecimal(this.getPoiRowAt(row).getHeight());
	}

	public BigDecimal getColWidthOf(int col) {
		return new BigDecimal(poiSheet.getColumnWidth(col));
	}

	public void showPageNumberAtCenterOfFooter() {
		this.poiSheet.getFooter().setCenter(
				HeaderFooter.page() + "/" + HeaderFooter.numPages());
	}

	public short getRowHeightAt(String refToCell) {
		CellReference ref = new CellReference(refToCell);
		Row row = getPoiRowAt(ref.getRow());
		return row.getHeight();
	}

	public Sheet getPoiSheet() {
		return this.poiSheet;
	}

	public void copyPrintSettingTo(ScSheet target) {

		PrintSetup targetSetting = target.getPoiSheet().getPrintSetup();
		PrintSetup templateSetting = this.getPoiSheet().getPrintSetup();

		targetSetting.setCopies(templateSetting.getCopies());
		targetSetting.setDraft(templateSetting.getDraft());
		targetSetting.setHResolution(templateSetting.getHResolution());
		targetSetting.setLandscape(templateSetting.getLandscape());
		targetSetting.setLeftToRight(templateSetting.getLeftToRight());
		targetSetting.setNoColor(templateSetting.getNoColor());
		targetSetting.setNoOrientation(templateSetting.getNoOrientation());
		targetSetting.setNotes(templateSetting.getNotes());
		targetSetting.setPageStart(templateSetting.getPageStart());
		targetSetting.setPaperSize(templateSetting.getPaperSize());
		targetSetting.setScale(templateSetting.getScale());
		targetSetting.setUsePage(templateSetting.getUsePage());
		targetSetting.setValidSettings(templateSetting.getValidSettings());
		targetSetting.setVResolution(templateSetting.getVResolution());
		targetSetting.setFitHeight(templateSetting.getFitHeight());
		targetSetting.setFitWidth(templateSetting.getFitWidth());

		target.getPoiSheet().setAutobreaks(this.getPoiSheet().getAutobreaks());

		targetSetting.setFooterMargin(templateSetting.getFooterMargin());
		targetSetting.setHeaderMargin(templateSetting.getHeaderMargin());

		copyTopBottomLeftRightMargine(target);

	}

	private void copyTopBottomLeftRightMargine(ScSheet target) {
		Sheet targetSheet = target.getPoiSheet();
		Sheet templateSheet = this.getPoiSheet();

		targetSheet.setMargin(HSSFSheet.TopMargin, templateSheet
				.getMargin(HSSFSheet.TopMargin));
		targetSheet.setMargin(HSSFSheet.BottomMargin, templateSheet
				.getMargin(HSSFSheet.BottomMargin));
		targetSheet.setMargin(HSSFSheet.LeftMargin, templateSheet
				.getMargin(HSSFSheet.LeftMargin));
		targetSheet.setMargin(HSSFSheet.RightMargin, templateSheet
				.getMargin(HSSFSheet.RightMargin));

	}

	public List<ScRange> getMergedRanges() {
		List<ScRange> ranges = new ArrayList<ScRange>();

		for (int i = 0; i < this.getPoiSheet().getNumMergedRegions(); i++) {
			ranges.add(ScRange
					.createFrom(this.getPoiSheet().getMergedRegion(i)));
		}

		return ranges;
	}

	public void copyMargedCellInfoTo(ScSheet target) {

		Sheet targetSheet = target.getPoiSheet();
		Sheet templateSheet = this.getPoiSheet();

		for (int i = 0; i < templateSheet.getNumMergedRegions(); i++) {
			targetSheet.addMergedRegion(templateSheet.getMergedRegion(i));
		}
	}

	public ScPaperSize getPaperSize() {
		return ScPaperSize.getFromPoiPaperIndex(this.getPoiSheet()
				.getPrintSetup().getPaperSize());
	}

	public String getPaperOrientation() {
		// TODO 安易なので考える
		return (this.getPoiSheet().getPrintSetup().getLandscape() ? "横" : "縦");
	}

	public ScWorkbook getWorkBook() {
		return wb;
	}

	public ScCell getCellAt(int rowPos, int colPos) {
		return getCellAt(new ScPos(rowPos, colPos));
	}

	public ScCell getCellAt(String pos) {
		return this.getCellAt(new ScPos(pos));
	}

	public ScCell getCellAt(ScPos pos) {
		ScCell c = ScCell.getAt(pos, this);
		return c;
	}

	public int incrementRow(int rowNumber) {
		currentRow += rowNumber;
		return currentRow;
	}

	public int getCurrentRow() {
		return currentRow;
	}

	public void setCurrentRow(int currentRow) {
		this.currentRow = currentRow;
	}

}
