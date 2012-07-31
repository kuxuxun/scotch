package com.github.kuxuxun.scotchtest.assertion.assertion;

import java.util.ArrayList;
import java.util.List;

import com.github.kuxuxun.scotch.excel.area.ScPos;

public class Positions extends In {

	private final List<ScPos> positions;

	private Positions(String[] poses) {
		super("", "");
		positions = new ArrayList<ScPos>();

		for (String each : poses) {
			positions.add(new ScPos(each));
		}
	}

	public static Positions in(String[] posisions) {
		return new Positions(posisions);
	}

	@Override
	public List<ScPos> getPositions() {
		return this.positions;
	}

}
