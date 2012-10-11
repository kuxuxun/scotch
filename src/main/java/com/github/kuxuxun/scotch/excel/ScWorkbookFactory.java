package com.github.kuxuxun.scotch.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ScWorkbookFactory {

	public static ScWorkbook createFromFile(File wbfile) throws IOException,
			InvalidFormatException {
		return create(new FileInputStream(wbfile));

	}

	public static ScWorkbook create(InputStream is) throws IOException,
			InvalidFormatException {

		return new ScWorkbook(WorkbookFactory.create(is));
	}

	/**
	 * ワークブックファイルの種類を表現します
	 */
	public enum WBFileKind {
		XLSM("xlsm"), XLSX("xlsx"), XML("xml"), XLS("xls");
		private final String extractor;

		private WBFileKind(String extractor) {
			this.extractor = extractor;
		}

		/**
		 * ファイルの拡張子からファイルの種類を判定します。
		 * <p>
		 * 判定できなかった場合はXLSを返却します。
		 * 
		 * e.g
		 * <ul>
		 * <li>/template/excel/foo.xls => XLS</li>
		 * <li>/template/excel/foo.XLS => XLS</li>
		 * <li>/template/excel/foo.xml => XML</li>
		 * <li>/template/excel/foo.bar=> XLS</li>
		 * </ul>
		 * 
		 * @param path
		 * @return
		 */
		public static WBFileKind getFromFile(File file) {
			return getFromFileFullPath(file.getAbsolutePath());
		}

		public static WBFileKind getFromFileFullPath(String path) {
			String e = getExtractorOf(path).toUpperCase();

			for (WBFileKind each : WBFileKind.values()) {
				if (each.extractor.equals(e)) {
					return each;
				}
			}

			return XLS;

		}

		private static String getExtractorOf(String path) {
			String fileName = extractFileNameFromPath(path);
			int p = fileName.lastIndexOf(".");
			if (p < 0) {
				return "";
			}

			return fileName.substring(p + 1);
		}

		private static String extractFileNameFromPath(String path) {
			if (n2b(path).length() == 0) {
				throw new IllegalArgumentException("ファイルを指定してください");
			}

			File f = new File(path);
			if (f.isDirectory()) {
				throw new IllegalArgumentException("ファイルを指定してください:" + path);
			}
			return f.getName();
		}
	}

	private static String n2b(String s) {
		if (s == null) {
			return "";
		}
		return s;
	}

}