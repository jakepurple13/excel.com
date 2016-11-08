package com.excel.excel.com;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMethods {

	final int NAME = 0;
	final int TIMEIN = 5;
	final int TIMEOUT = 6;

	final int DAY = 1;
	final int STARTTIME = 2;
	final int ENDTIME = 3;

	HashMap<String, Student> hmss;
	ArrayList<Exception> badThings;
	ArrayList<String> timing;
	ArrayList<Date> starting;
	ArrayList<Date> ending;
	ArrayList<WorkTime> workTime;

	String timeReadInName = "schedules_on_kronos_201670.xls";
	String realReadInName = "EmployeeTimeDetail_PayPeriodEnd10-15-16.xlsx";

	ArrayList<String> name = new ArrayList<String>();

	public ExcelMethods() {

		hmss = new HashMap<String, Student>();
		badThings = new ArrayList<Exception>();
		timing = new ArrayList<String>();
		starting = new ArrayList<Date>();
		ending = new ArrayList<Date>();
		workTime = new ArrayList<WorkTime>();

		// READS SCHEDULE WORKBOOK
		// suspectedTimeReadIn(timeReadInName);
		timeInRead();
		/*
		 * System.out.println(hmss.values()); System.out.println(hmss.keySet());
		 */
		for (Student s : hmss.values()) {
			System.err.println(s);
		}

		// System.out.println("------------");

		// -----------------------------------

		// READS WORKBOOK THAT HAS EXACT TIME DETAIL
		// actualReadIn(realReadInName);

		// System.err.println("------------");

		// ------------------------------

		// CREATES WORKBOOK AND ADDS DATA TO IT
		// writeOut();
	}

	public void timeInRead() {
		try {
			suspectedTimeReadIn(timeReadInName);
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void suspectedTimeReadIn(String fileName)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook wb1 = WorkbookFactory.create(new File(fileName));

		Sheet sheet1 = wb1.getSheet("schedule matches google drive");

		for (int i = 1; i < sheet1.getLastRowNum(); i++) {

			Row r = sheet1.getRow(i); // the row we are on
			String day; // the day
			String named; // the name of the employee
			String startTime; // the time the employee is going to start working
			String endTime; // the time the employee is going to end working
			DateTimeFormatter parseFormat; // will parse a date into a LocalTime
											// format
			LocalTime start; // the start time in LocalTime format
			LocalTime end; // the end time in LocalTime format
			Cell two; // the second column which should be the start time
			Cell three; // the third column which should be the end time

			named = r.getCell(NAME).toString();

			if (named.equals("")) {
				break;
			}

			day = r.getCell(DAY).toString();

			two = r.getCell(STARTTIME);
			three = r.getCell(ENDTIME);

			if (DateUtil.isCellDateFormatted(two)) {
				startTime = new DataFormatter().formatCellValue(two);
			} else {
				startTime = String.valueOf((int) (two.getNumericCellValue()));
			}

			if (DateUtil.isCellDateFormatted(three)) {
				endTime = new DataFormatter().formatCellValue(three);
			} else {
				endTime = String.valueOf((int) (three.getNumericCellValue()));
			}

			parseFormat = new DateTimeFormatterBuilder().appendPattern("h:mm:ss a").toFormatter();

			start = LocalTime.parse(startTime, parseFormat);
			end = LocalTime.parse(endTime, parseFormat);

			if (hmss.containsKey(named)) {
				hmss.get(named).addTime(day, start, end);
			} else {
				Student s = new Student(named);
				hmss.put(named, s);
				hmss.get(named).addTime(day, start, end);
			}
		}
	}

	public void actualIn(String fileName) {
		try {
			actualReadIn(fileName);
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void actualReadIn(String fileName)
			throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException {
		Workbook wb = WorkbookFactory.create(new File(fileName));

		Sheet sheet = wb.getSheet("Spans");

		for (int rows = 5; rows < sheet.getLastRowNum(); rows++) {

			Date in = null;
			Date out = null;
			String named;

			Row r = sheet.getRow(rows);

			Cell a = r.getCell(TIMEIN);
			Cell c = r.getCell(TIMEOUT);

			named = r.getCell(NAME).toString().trim();

			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm");

			in = sdf.parse(a.toString());
			out = sdf.parse(c.toString());

			System.err.println(named + "\t" + in + "\t" + out);

			if (hmss.containsKey(named)) {

				Student s = hmss.get(named);

				System.out.println(s.toString());

				System.out.println(s.checkTime(in) + "\t" + s.checkTime(out));
				timing.add("Start change: " + s.checkTime(in) + " minutes" + "\tEnd change: " + s.checkTime(out)
						+ " minutes");

				workTime.add(new WorkTime(s.checkTime(in), s.checkTime(out)));

				starting.add(in);
				ending.add(out);

				name.add(named);
			} else {
				new MissingPersonException(named, badThings);
				// badThings.add(new MissingPersonException(named));
			}

		}
	}

	public void outWrite(String fileName) {
		try {
			writeOut(fileName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void writeOut(String fileName) throws IOException {
		// Workbook flagged = new HSSFWorkbook();
		Workbook flagged = new XSSFWorkbook();
		// CreationHelper createHelper = flagged.getCreationHelper();
		Sheet flagShip = flagged.createSheet("People who are bad");
		ArrayList<Student> als = new ArrayList<Student>(hmss.values());

		Row categories = flagShip.createRow(0);

		Cell cat = categories.createCell(NAME);
		cat.setCellValue("Name");

		cat = categories.createCell(1);
		cat.setCellValue("Days of Work");

		cat = categories.createCell(2);
		cat.setCellValue("Starting Time");

		cat = categories.createCell(3);
		cat.setCellValue("Ending Time");

		cat = categories.createCell(4);
		cat.setCellValue("In Difference");

		cat = categories.createCell(5);
		cat.setCellValue("Out Difference");

		int j = 0;
		int count = 0;
		for (int i = 0; i < als.size(); i++) {

			// System.out.println(als.get(i));
			ArrayList<Times> alt = als.get(i).getTime();

			System.err.println(count + "\tals size: " + als.size());

			// Create a row and put some cells in it. Rows are 0 based.
			// Row row = flagShip.createRow(count);

			// Create a cell and put a value in it.
			// Cell cell = row.createCell(NAME);
			// cell.setCellValue(als.get(i).name);

			for (j = 0; j < alt.size(); j++) {

				count++;

				System.out.println("Name: " + als.get(i).name + "\tData: " + alt.get(j) + "\tsize: " + alt.size()
						+ "\ti: " + i + "\tj: " + j + "\tcount: " + count);

				Row row = flagShip.createRow(count);
				if (j == 0) {
					Cell nameCell = row.createCell(NAME);
					nameCell.setCellValue(als.get(i).name);
				}

				Cell daysOfWorkCell = row.createCell(1);
				daysOfWorkCell.setCellValue(alt.get(j).toString());

				Cell startingTimeCell = row.createCell(2);
				startingTimeCell.setCellValue(starting.get(count - 1).toString());

				Cell endingTimeCell = row.createCell(3);
				endingTimeCell.setCellValue(ending.get(count - 1).toString());

				Cell inDifferenceCell = row.createCell(4);
				inDifferenceCell.setCellValue(workTime.get(count - 1).getIn());

				Cell outDifferenceCell = row.createCell(5);
				outDifferenceCell.setCellValue(workTime.get(count - 1).getOut());

				// count++;

			}

		}

		// auto size columns used to make things look nice
		flagShip.autoSizeColumn(0);
		flagShip.autoSizeColumn(1);
		flagShip.autoSizeColumn(2);
		flagShip.autoSizeColumn(3);
		flagShip.autoSizeColumn(4);
		flagShip.autoSizeColumn(5);

		SheetConditionalFormatting sheetCF = flagShip.getSheetConditionalFormatting();
		ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "10");
		PatternFormatting fill1 = rule1.createPatternFormatting();
		fill1.setFillBackgroundColor(IndexedColors.ROSE.index);
		fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] regions = { CellRangeAddress.valueOf("E2:F82") };

		sheetCF.addConditionalFormatting(regions, rule1);

		// -------------------------------

		Sheet allTime = flagged.createSheet("Times");

		for (int i = 0; i < starting.size(); i++) {

			try {

				// Create a row and put some cells in it. Rows are 0 based.
				Row row = allTime.createRow(i);
				// Create a cell and put a value in it.
				Cell nameCell = row.createCell(0);
				nameCell.setCellValue(name.get(i));

				Cell startingCell = row.createCell(1);
				startingCell.setCellValue(starting.get(i).toString());

				Cell endingCell = row.createCell(2);
				endingCell.setCellValue(ending.get(i).toString());

				Cell workTimeInCell = row.createCell(3);
				workTimeInCell.setCellValue(workTime.get(i).getIn());

				Cell workTimeOutCell = row.createCell(4);
				workTimeOutCell.setCellValue(workTime.get(i).getOut());

				Cell tester = row.createCell(5);

				String named = name.get(i);
				System.out.println(named);
				Date d = starting.get(i);
				System.err.println(d.toString());
				Student s = hmss.get(name.get(i));
				System.err.println(s);
				Times t = s.getDays(d);
				System.out.println(t.toString());

				tester.setCellValue(t.toString());

			} catch (NullPointerException e) {
				new MissingPersonException(name.get(i) + "\t" + e.toString(), badThings);
			}

		}

		SheetConditionalFormatting sheetAT = allTime.getSheetConditionalFormatting();
		ConditionalFormattingRule rule2 = sheetAT.createConditionalFormattingRule(ComparisonOperator.GT, "10");
		PatternFormatting fill2 = rule2.createPatternFormatting();
		fill2.setFillBackgroundColor(IndexedColors.RED.index);
		fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] regions1 = { CellRangeAddress.valueOf("D1:E1000") };

		sheetAT.addConditionalFormatting(regions1, rule2);

		allTime.autoSizeColumn(0);
		allTime.autoSizeColumn(1);
		allTime.autoSizeColumn(2);
		allTime.autoSizeColumn(3);
		allTime.autoSizeColumn(4);

		Sheet badShip = flagged.createSheet("Exceptions");

		for (int i = 0; i < badThings.size(); i++) {

			// Create a row and put some cells in it. Rows are 0 based.
			Row row = badShip.createRow(i);
			// Create a cell and put a value in it.
			Cell cell = row.createCell(0);
			cell.setCellValue(badThings.get(i).getMessage());
			// cell = row.createCell(1);
			// cell.setCellValue(((MissingPersonException)
			// badThings.get(i)).errorMessage);

		}

		badShip.autoSizeColumn(0);
		badShip.autoSizeColumn(1);

		// File desktop = new File(System.getProperty("user.home"), "Desktop");
		// Write the output to a file
		/*
		 * FileOutputStream fileOut = new FileOutputStream(
		 * desktop.getAbsolutePath() + "/captainslog.xlsx");
		 */
		FileOutputStream fileOut = new FileOutputStream(fileName);
		flagged.write(fileOut);
		fileOut.close();
		flagged.close();
	}

}
