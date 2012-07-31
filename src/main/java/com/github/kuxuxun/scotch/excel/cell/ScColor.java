package com.github.kuxuxun.scotch.excel.cell;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.IndexedColors;

//FIXME index 2は未対応
public enum ScColor {

	BLACK(IndexedColors.BLACK.getIndex()), BROWN(IndexedColors.BROWN.getIndex()), OLIVE_GREEN(
			IndexedColors.OLIVE_GREEN.getIndex()), DARK_GREEN(
			IndexedColors.DARK_GREEN.getIndex()), DARK_TEAL(
			IndexedColors.DARK_TEAL.getIndex()), DARK_BLUE(
			IndexedColors.DARK_BLUE.getIndex()), INDIGO(IndexedColors.INDIGO
			.getIndex()), GREY_80_PERCENT(IndexedColors.GREY_80_PERCENT
			.getIndex()), ORANGE(IndexedColors.ORANGE.getIndex()), DARK_YELLOW(
			IndexedColors.DARK_YELLOW.getIndex()), GREEN(IndexedColors.GREEN
			.getIndex()), TEAL(IndexedColors.TEAL.getIndex()), BLUE(
			IndexedColors.BLUE.getIndex()), BLUE_GREY(IndexedColors.BLUE_GREY
			.getIndex()), GREY_50_PERCENT(IndexedColors.GREY_50_PERCENT
			.getIndex()), RED(IndexedColors.RED.getIndex()), LIGHT_ORANGE(
			IndexedColors.LIGHT_ORANGE.getIndex()), LIME(IndexedColors.LIME
			.getIndex()), SEA_GREEN(IndexedColors.SEA_GREEN.getIndex()), AQUA(
			IndexedColors.AQUA.getIndex()), LIGHT_BLUE(IndexedColors.LIGHT_BLUE
			.getIndex()), VIOLET(IndexedColors.VIOLET.getIndex()), GREY_40_PERCENT(
			IndexedColors.GREY_40_PERCENT.getIndex()), PINK(IndexedColors.PINK
			.getIndex()), GOLD(IndexedColors.GOLD.getIndex()), YELLOW(
			IndexedColors.YELLOW.getIndex()), BRIGHT_GREEN(
			IndexedColors.BRIGHT_GREEN.getIndex()), TURQUOISE(
			IndexedColors.TURQUOISE.getIndex()), DARK_RED(
			IndexedColors.DARK_RED.getIndex()), SKY_BLUE(IndexedColors.SKY_BLUE
			.getIndex()), PLUM(IndexedColors.PLUM.getIndex()), GREY_25_PERCENT(
			IndexedColors.GREY_25_PERCENT.getIndex()), ROSE(IndexedColors.ROSE
			.getIndex()), LIGHT_YELLOW(IndexedColors.LIGHT_YELLOW.getIndex()), LIGHT_GREEN(
			IndexedColors.LIGHT_GREEN.getIndex()), LIGHT_TURQUOISE(
			IndexedColors.LIGHT_TURQUOISE.getIndex()), PALE_BLUE(
			IndexedColors.PALE_BLUE.getIndex()), LAVENDER(
			IndexedColors.LAVENDER.getIndex()), WHITE(IndexedColors.WHITE
			.getIndex()), CORNFLOWER_BLUE(IndexedColors.CORNFLOWER_BLUE
			.getIndex()), LEMON_CHIFFON(IndexedColors.LEMON_CHIFFON.getIndex()), MAROON(
			IndexedColors.MAROON.getIndex()), ORCHID(IndexedColors.ORCHID
			.getIndex()), CORAL(IndexedColors.CORAL.getIndex()), ROYAL_BLUE(
			IndexedColors.ROYAL_BLUE.getIndex()), LIGHT_CORNFLOWER_BLUE(
			IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex()), TAN(
			IndexedColors.TAN.getIndex()), NON_OR_AUTO(IndexedColors.AUTOMATIC
			.getIndex());

	private static Map<Short, ScColor> indexTable;
	static {
		indexTable = new HashMap<Short, ScColor>();
		for (ScColor each : ScColor.values()) {
			indexTable.put(each.getIndex(), each);
		}

	}

	private short index;

	private ScColor(short i) {
		this.index = i;
	}

	public short getIndex() {
		return index;
	}

	public static ScColor getFromColorIndex(short index) {
		ScColor c = indexTable.get(index);
		if (c != null) {
			return c;
		}
		return NON_OR_AUTO;
	}

}
