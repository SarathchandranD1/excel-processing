package com.excel.util;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class Utility {

	public static String getFileExtension(String fileName) {
		return fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
	}

	public static Date getDateWithoutTimeUsingCalendar() {
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.HOUR_OF_DAY, 5);
		calendar.set(Calendar.MINUTE, 30);
		calendar.set(Calendar.SECOND, 0);
		calendar.set(Calendar.MILLISECOND, 0);
		return calendar.getTime();
	}

	public static List<Integer> getDaysByWeekMonthYear(int weekIndex, int month, int year) {
		List<List<Integer>> weeks = new ArrayList<>();
		LocalDate firstDayOfMonth = LocalDate.of(year, month, 1);
		LocalDate firstWeekendOfMonth = firstDayOfMonth.with(TemporalAdjusters.firstInMonth(DayOfWeek.SUNDAY));
		long firstWeekGap = ChronoUnit.DAYS.between(firstDayOfMonth, firstWeekendOfMonth);
		List<Integer> daysInWeek = buildDaysInWeek(1, month, year);
		weeks.add(daysInWeek);
		for (int i = ((int) firstWeekGap + 1); i <= 31; i = i + 7) {
			daysInWeek = buildDaysInWeek(i, month, year);
			weeks.add(daysInWeek);

		}
		return weekIndex >= 5 ? weeks.get(weekIndex - 1).stream().filter(x -> x > 7).collect(Collectors.toList())
				: weeks.get(weekIndex - 1);
	}

	private static List<Integer> buildDaysInWeek(int date, int month, int year) {
		List<Integer> daysInWeek = new ArrayList<>();
		LocalDate startDate = LocalDate.of(year, month, date);
		daysInWeek.add(startDate.getDayOfMonth());
		LocalDate endOfWeek = startDate;
		while (endOfWeek.getDayOfWeek() != DayOfWeek.SATURDAY) {
			endOfWeek = endOfWeek.plusDays(1);
			daysInWeek.add(endOfWeek.getDayOfMonth());
		}
		return daysInWeek;
	}
	
	public static LocalDate extractDateFromExcel(String fileName) {
		String extract = fileName.substring(fileName.length()-13,fileName.length()-5);
		int year = Integer.parseInt(extract.substring(0,4));
		int month = Integer.parseInt(extract.substring(4,6));
		int dayOfMonth = Integer.parseInt(extract.substring(6,8));
		return LocalDate.of(year, month, dayOfMonth);
	}
}
