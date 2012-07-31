package com.github.kuxuxun.scotch.util;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class RefrectionUtil {
	public static <T> Method getMethodAsAccessible(String methodName,
			Class<T> clazz) {

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

	private static Method methodIsCorrectOrNull(String methodName, Method each) {
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

	public static void setValueToField(Object reciever, String fieldName,
			Object arg) throws IllegalArgumentException, IllegalAccessException {

		Field f = getFieldAsAccessible(fieldName, reciever.getClass());

		f.set(reciever, arg);

	}

	public static Object callMethod(Object reciever, String methodName,
			Object... args) throws Throwable {

		try {
			Method m = getMethodAsAccessible(methodName, reciever.getClass());
			return m.invoke(reciever, args);
		} catch (InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

}
