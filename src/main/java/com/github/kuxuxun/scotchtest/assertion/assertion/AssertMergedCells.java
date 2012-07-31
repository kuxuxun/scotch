package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.area.ScRange;
import com.github.kuxuxun.scotch.excel.cell.ScCell;

public class AssertMergedCells {

	private final ScSheet sheet;

	private AssertMergedCells(ScSheet report) {
		this.sheet = report;
	}

	public static AssertMergedCells that(ScSheet file) {
		return new AssertMergedCells(file);
	}

	public void cellsAreMerged(In designatedRange)
			throws FileNotFoundException, IOException {
		List<ScRange> ranges = sheet.getMergedRanges();

		ScRange correspondingRange = ScRange.selectWhichContainOrNull(
				designatedRange.getPositions().get(0), ranges);

		if (correspondingRange == null) {
			throw new IllegalArgumentException("指定のセルを含む、結合セルがありません："
					+ designatedRange.getPositions().get(0).toString());
		}

		final Set<String> designatedRefs = new HashSet<String>();
		for (ScPos each : designatedRange.getPositions()) {
			designatedRefs.add(each.toReference());
		}

		In mergedCellAreaOfSheet = In.fromScRange(correspondingRange);

		new AssertIn(sheet, mergedCellAreaOfSheet) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				boolean result = designatedRefs.contains(cell.getPos()
						.toReference());

				assertTrue(cell.getPos().toReference() + "もマージ範囲に含まれている: ",
						result);
			}
		}.doAssert();

	}

	public void cellsAreNotMerged(In designatedRange)
			throws FileNotFoundException, IOException {
		final List<ScRange> ranges = sheet.getMergedRanges();

		new AssertIn(sheet, designatedRange) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				ScRange containingRange = ScRange.selectWhichContainOrNull(cell
						.getPos(), ranges);

				assertNull(cell.getPos().toReference() + "がマージ範囲に含まれていない: ",
						containingRange);
			}
		}.doAssert();

	}

}
