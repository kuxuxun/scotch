package com.github.kuxuxun.scotch.excel.area;

import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

public class ScRange extends ScArea {

	private final ScPos topLeft;
	private final ScPos bottomRight;

	public static ScRange selectWhichContainOrNull(ScPos pos,
			List<ScRange> ranges) {
		for (ScRange each : ranges) {
			if (each.contains(pos)) {
				return each;
			}
		}

		return null;
	}

	public boolean contains(ScPos pos) {

		if ((topLeft.getRow() <= pos.getRow() && pos.getRow() <= bottomRight
				.getRow())
				&&

				(topLeft.getCol() <= pos.getCol() && pos.getCol() <= bottomRight
						.getCol())) {
			return true;
		}

		return false;

	}

	public static ScRange createFrom(CellRangeAddress address) {

		ScPos topLeft = new ScPos(address.getFirstRow(), address
				.getFirstColumn());

		ScPos bottomRight = new ScPos(address.getLastRow(), address
				.getLastColumn());

		ScRange r = new ScRange(topLeft, bottomRight);

		return r;
	}

	public ScRange(ScPos topLeft, ScPos bottomRight) {
		this.topLeft = topLeft;
		this.bottomRight = bottomRight;
	}

	public static ScRange in(int topRow, int leftCol, int bottomRow,
			int rightCol) {
		return new ScRange(new ScPos(topRow, leftCol), new ScPos(bottomRow,
				rightCol));
	}

	@Override
	public String toReference() {
		CellReference topLeftRef = new CellReference(topLeft.getRow(), topLeft
				.getCol(), false, false);
		CellReference bottomRightRef = new CellReference(bottomRight.getRow(),
				bottomRight.getCol(), false, false);

		return topLeftRef.formatAsString() + ":"
				+ bottomRightRef.formatAsString();
	}

	public ScPos getTopLeft() {
		return topLeft;
	}

	public ScPos getBottomRight() {
		return bottomRight;
	}

	@Override
	public int getBottomRow() {
		return this.getBottomRight().getRow();
	}

	@Override
	public int getTopRow() {
		return this.getTopLeft().getRow();
	}

	@Override
	public int getLeftCol() {
		return this.getTopLeft().getCol();
	}

	@Override
	public int getRightCol() {
		return this.bottomRight.getCol();
	}

}
