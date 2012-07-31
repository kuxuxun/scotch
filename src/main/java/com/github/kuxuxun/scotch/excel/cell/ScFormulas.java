package com.github.kuxuxun.scotch.excel.cell;

import java.util.ArrayList;
import java.util.List;

import com.github.kuxuxun.scotch.excel.area.ScPos;


public class ScFormulas {

	/**
	 * 列(縦方向)のsum関数を作成します。
	 * 
	 * @param startRow
	 * @param endRow
	 * @param col
	 * @return
	 */
	public static String createVerticalSum(int startRow, int endRow, int col) {
		return createsequentialsum(new ScPos(startRow, col), new ScPos(endRow, col));

	}

	/**
	 * 列(縦方向)のsum関数のListを作成します
	 * 
	 * @param startRow
	 * @param endRow
	 * @param startCol
	 * @param numberOfRol
	 * @return
	 */
	public static List<String> createSeveralVerticalSum(int startRow,
			int endRow, int startCol, int numberOfCol) {

		List<String> result = new ArrayList<String>();

		for (int colOperand = 0; colOperand < numberOfCol; colOperand++) {
			int col = startCol + colOperand;
			result.add(createsequentialsum(new ScPos(startRow, col), new ScPos(
					endRow, col)));
		}

		return result;

	}

	/**
	 * 行(横方向)のsum関数のListを作成します
	 * 
	 * @param startcol
	 * @param endCol
	 * @param row
	 * @return
	 */
	public static List<String> createSeveralHorizontalSum(int startCol,
			int endCol, int startRow, int numberOfRow) {

		List<String> result = new ArrayList<String>();

		for (int rowOperand = 0; rowOperand < numberOfRow; rowOperand++) {
			int row = startRow + rowOperand;
			result.add(createsequentialsum(new ScPos(row, startCol), new ScPos(row,
					endCol)));
		}

		return result;

	}

	/**
	 * 行(横方向)のsum関数を作成します
	 * 
	 * @param startcol
	 * @param endcol
	 * @param row
	 * @return
	 */
	public static String createHorizontalSum(int startcol, int endcol, int row) {
		return createsequentialsum(new ScPos(row, startcol), new ScPos(row, endcol));

	}

	private static String createsequentialsum(ScPos startpos, ScPos endpos) {
		String result = "";
		result += "sum(";
		result += startpos.toReference();
		result += ":";
		result += endpos.toReference();

		result += ")";

		return result;
	}

	/**
	 * 渡された引数のセルの合計関数を返却します
	 * 
	 * @param sellPositions
	 * @return
	 */
	public static String createSequentialAdding(ScPos... sellPositions) {
		String result = "";

		for (ScPos cellPos : sellPositions) {
			if (result.length() != 0) {
				result += "+";
			}

			result += cellPos.toReference();

		}

		return result;
	}

	/**
	 * 複数の水平方向の引き算式を作成します。
	 * 
	 * @param startcol
	 * @param endcol
	 * @param row
	 * @return
	 */
	public static List<String> createSeveralHorizontalSubtraction(int lhsCol,
			int rhsCol, int startRow, int numberOfRow) {
		List<String> result = new ArrayList<String>();

		for (int rowOperand = 0; rowOperand < numberOfRow; rowOperand++) {
			int row = startRow + rowOperand;
			ScPos lhs = new ScPos(row, lhsCol);
			ScPos rhs = new ScPos(row, rhsCol);

			result.add(lhs.toReference() + "-" + rhs.toReference());
		}

		return result;

	}

	/**
	 * 比率をあらわすexcelの式を返却します。 除数が0の場合は、0とする式が返却されます。
	 * <p>
	 * 比率を計算する式は"ROUND(A1/B1*100,30)"のように31桁目を四捨五入する式となります。
	 * ROUNDを使用せず"A1/B1*100"という式にすると式が#Value!と評価されてしまうため、
	 * Excelの数値の小数点以下最大表示桁30桁でROUNDを行います。
	 * 
	 * @param divident
	 * @param divisor
	 * @return
	 */
	public static String createRateFunction(ScPos divident, ScPos divisor) {

		return "IF(" + divisor.toReference() + "=0,0,ROUND("
				+ divident.toReference() + "/" + divisor.toReference()
				+ "*100,30))";
	}
}
