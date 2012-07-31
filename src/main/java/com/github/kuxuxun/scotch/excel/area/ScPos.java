package com.github.kuxuxun.scotch.excel.area;

import org.apache.poi.ss.util.CellReference;

public class ScPos extends ScArea {

	private final int row;
	private final int col;

	public ScPos(int row, int col) {
		this.row = row;
		this.col = col;
	}

	public boolean isLeftEdge() {
		if (col == 0) {
			return true;
		}
		return false;
	}

	public boolean isTopEdge() {
		if (row == 0) {
			return true;
		}
		return false;
	}

	public ScPos(String refernce) {
		CellReference ref = new CellReference(refernce);
		this.row = ref.getRow();
		this.col = ref.getCol();
	}

	public int getRow() {
		return row;
	}

	public int getCol() {
		return col;
	}

	public ScPos moveInCol(int num) {

		if ((this.col + num) < 0) {
			return this;
		}

		return new ScPos(row, col + num);
	}

	public ScPos moveInRow(int num) {
		if ((this.row + num) < 0) {
			return this;
		}

		return new ScPos(row + num, col);
	}

	/**
	 * セルの座標から参照文字列を取得します。
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	@Override
	public String toReference() {
		CellReference ref = new CellReference(row, col, false, false);
		return ref.formatAsString();
	}

	@Override
	public String toString() {
		return toReference();
	}

	@Override
	public int getBottomRow() {
		return row;
	}

	@Override
	public int getTopRow() {
		return row;
	}

	@Override
	public int getLeftCol() {
		return col;
	}

	@Override
	public int getRightCol() {
		return col;
	}

	/**
	 * セルリファレンスがさす行を取得します
	 * 
	 * @param cellrefference
	 * @return
	 */
	public static int rowNumberOf(String cellrefference) {
		CellReference ref = new CellReference(cellrefference);
		return ref.getRow();
	}

	/**
	 * セルリファレンスがさす列を取得します
	 * 
	 * @param cellrefference
	 * @return
	 */
	public static int colNumberOf(String cellRefference) {
		CellReference ref = new CellReference(cellRefference);
		return ref.getCol();
	}

	@Override
	public boolean equals(Object o) {
		if (o == this) {
			return true;
		}
		if (!(o instanceof ScPos)) {
			return false;
		}
		ScPos that = (ScPos) o;

		if ((this.col == that.col) && (this.row == that.row)) {
			return true;
		}

		return false;

	}

	@Override
	public int hashCode() {
		int r = this.col;
		r += col * 31;
		return r;
	}

}
