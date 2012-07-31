package com.github.kuxuxun.commonutil.ut;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class TestHelper {

    Calendar cal = Calendar.getInstance();

	public static Date date(String s) throws ParseException {
		String[] p = s.split("/");
		// TODO if文多くてきもい
		if (p.length == 3) {
			return new SimpleDateFormat("yyyy/MM/dd").parse(s);
		} else if (p.length == 2) {
			return new SimpleDateFormat("yyyy/MM").parse(s);
		} else {
			if (s.length() == 8) {
				return new SimpleDateFormat("yyyyMMdd").parse(s);
			} else {
				return new SimpleDateFormat("yyyyMM").parse(s);
			}
		}

	}

	// @Ignore
	// @Test
	// public void helperReturnsCorrectNumber() throws Exception {
	// assertEquals(0, new BigDecimal("12.34").compareTo(decimal(12.34)));
	// assertEquals(0, new BigDecimal("12.341111")
	// .compareTo(decimal(12.341111)));
	// assertEquals(0, new BigDecimal("0.34").compareTo(decimal(0.34)));
	// assertEquals(0, new BigDecimal("12").compareTo(decimal(12)));
	// }
	//
	// @Ignore
	// @Test
	// public void helperReturnsCorrectDate() throws Exception {
	// assertEquals("2010/01", format(january(2010)));
	// assertEquals("2010/02", format(february(2010)));
	// assertEquals("2010/03", format(march(2010)));
	// assertEquals("2010/04", format(april(2010)));
	// assertEquals("2010/05", format(may(2010)));
	// assertEquals("2010/06", format(june(2010)));
	// assertEquals("2010/07", format(july(2010)));
	// assertEquals("2010/08", format(august(2010)));
	// assertEquals("2010/09", format(september(2010)));
	// assertEquals("2010/10", format(october(2010)));
	// assertEquals("2010/11", format(november(2010)));
	// assertEquals("2010/12", format(december(2010)));
	//
	// assertEquals("2011/04", format(april(2011)));
	//
	// }

	public Date january(int year) {
		setMonthFirst(year, 0);
		return cal.getTime();
	}

	public Date february(int year) {
		setMonthFirst(year, 1);
		return cal.getTime();
	}

	public Date march(int year) {
		setMonthFirst(year, 2);

		return cal.getTime();
	}

	public Date april(int year) {
		setMonthFirst(year, 3);

		return cal.getTime();
	}

	public Date may(int year) {
		setMonthFirst(year, 4);

		return cal.getTime();
	}

	public Date june(int year) {
		setMonthFirst(year, 5);

		return cal.getTime();
	}

	public Date july(int year) {
		setMonthFirst(year, 6);

		return cal.getTime();
	}

	public Date august(int year) {
		setMonthFirst(year, 7);
		return cal.getTime();
	}

	public Date september(int year) {
		setMonthFirst(year, 8);
		return cal.getTime();
	}

	public Date october(int year) {
		setMonthFirst(year, 9);
		return cal.getTime();
	}

	public Date november(int year) {
		setMonthFirst(year, 10);
		return cal.getTime();
	}

	public Date december(int year) {
		setMonthFirst(year, 11);
		return cal.getTime();
	}

	private void setMonthFirst(int year, int month) {
		cal.set(Calendar.YEAR, year);
		cal.set(Calendar.MONTH, month);
		cal.set(Calendar.DAY_OF_MONTH, 1);
	}

	public String format(Date date) {
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM");
		return formatter.format(date);
	}

	public static BigDecimal decimal(String val) {
		return new BigDecimal(val.toString());
	}

	public static BigDecimal decimal(Number val) {
		return new BigDecimal(val.toString());
	}

	public List<Date> firstHalfOf(int year) {
		List<Date> result = new ArrayList<Date>();
		result.add(april(year));
		result.add(may(year));
		result.add(june(year));
		result.add(july(year));
		result.add(august(year));
		result.add(september(year));
		return result;
	}

	public List<Date> secondHalfOf(int year) {
		List<Date> result = new ArrayList<Date>();
		result.add(october(year));
		result.add(november(year));
		result.add(december(year));
		result.add(january(year + 1));
		result.add(february(year + 1));
		result.add(march(year + 1));
		return result;
	}

	public static String rq(String src) {
		return src.replaceAll("\"", "");
	}

	public static boolean dateCotains(String target, List<Date> dates) {
		DateFormat f = new SimpleDateFormat("yyyyMM");
		for (Date each : dates) {
			if (f.format(each).equals(target)) {
				return true;
			}
		}
		return false;

	}

	public <T> Method getMethodAsAccessible(String methodName, Class<T> clazz) {

		for (Method each : clazz.getDeclaredMethods()) {

			Method result = methodIsCorrectOrNull(methodName, each);
			if (result != null)
				return result;
		}

		for (Method each : clazz.getSuperclass().getDeclaredMethods()) {
			Method result = methodIsCorrectOrNull(methodName, each);
			if (result != null)
				return result;
		}

		throw new IllegalArgumentException("There is no method named "
				+ methodName + " in class " + clazz);

	}

	private Method methodIsCorrectOrNull(String methodName, Method each) {
		if (each.getName().equals(methodName)) {
			each.setAccessible(true);
			return each;
		}
		return null;
	}

	public static <T> Field getFieldAsAccessible(String fieldName,
			Class<T> clazz) {

		for (Field each : clazz.getDeclaredFields()) {
			Field result = isCorrectFieldOrNull(fieldName, each);
			if (result != null) {
				return result;
			}
		}

		for (Field each : clazz.getSuperclass().getDeclaredFields()) {

			Field result = isCorrectFieldOrNull(fieldName, each);
			if (result != null) {
				return result;
			}
		}

		throw new IllegalArgumentException("There is no field named  "
				+ fieldName + " in class " + clazz);

	}

	private static Field isCorrectFieldOrNull(String fieldName, Field each) {
		if (each.getName().equals(fieldName)) {
			each.setAccessible(true);
			return each;
		}
		return null;
	}

	public static void setValueToFieldOf(Object reciever, String fieldName,
			Object arg) throws IllegalArgumentException, IllegalAccessException {

		Field f = getFieldAsAccessible(fieldName, reciever.getClass());

		f.set(reciever, arg);

	}

	public Object callMethodOf(Object reciever, String methodName,
			Object... args) throws Throwable {

		try {
			Method m = getMethodAsAccessible(methodName, reciever.getClass());
			return m.invoke(reciever, args);
		} catch (InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

	public void assertCompareEquals(BigDecimal e, BigDecimal t) {
		assertCompareEquals("", e, t);
	}

	public void assertCompareEquals(String desc, BigDecimal e, BigDecimal t) {
		if (e.compareTo(t) == 0) {
			assertTrue(desc, e.compareTo(t) == 0);
		} else {
			assertEquals(desc, e, t);
		}

	}

}
