package com.github.kuxuxun.commonutil.ut;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class TestOutput {

	public static FileOutputStream getAsOutputStream(String fileName) {

		FileOutputStream fos = null;
		try {
			return new FileOutputStream("/Temp/" + fileName);
		} catch (FileNotFoundException ex) {
			new IllegalArgumentException("/Temp Not found");
		}

		return fos;

	}

}
