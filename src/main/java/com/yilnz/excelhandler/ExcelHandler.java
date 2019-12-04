package com.yilnz.excelhandler;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class ExcelHandler {

	private static final String EXCEL_XLS = "xls";
	private static final String EXCEL_XLSX = "xlsx";

	private int handlerSeq = 1;

	public static Workbook getWorkbook(File file) throws IOException {
		Workbook wb = null;
		FileInputStream in = new FileInputStream(file);
		if (file.getName().endsWith(EXCEL_XLS)) {     //Excel&nbsp;2003
			wb = new HSSFWorkbook(in);
		} else if (file.getName().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
			wb = new XSSFWorkbook(in);
		}
		return wb;
	}

	private void doHandleExcelSeq(Row startRow, Row endRow) {
		for (int i = startRow.getRowNum(); i <= endRow.getRowNum(); i++) {
			final Cell cell = startRow.getCell(0);
			if (cell != null) {
				final Row row = startRow.getSheet().getRow(i);
				row.removeCell(cell);
			}
		}


		Cell cell = startRow.getCell(0);
		if (cell == null) {
			cell = startRow.createCell(0);
		}
		final CellStyle cellStyle = cell.getCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cell.setCellValue(handlerSeq++);

		//isMerged(startRow, endRow);
		final Sheet sheet = startRow.getSheet();
		try {
			sheet.addMergedRegion(new CellRangeAddress(startRow.getRowNum(), endRow.getRowNum(), 0, 0));
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
			//isMerged(startRow, endRow);
		}

		//final CellRangeAddress mergedRegion = sheet.getMergedRegion(region);
	}

	private void isMerged(Row start, Row end){
		final Sheet sheet = start.getSheet();
		//final List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
		for (int i = 0; i < sheet.getNumMergedRegions();i++) {
			final CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
			final int firstRow = mergedRegion.getFirstRow();
			final int lastRow = mergedRegion.getLastRow();
			if(!(firstRow > end.getRowNum() || lastRow < start.getRowNum())){
				sheet.removeMergedRegion(i);
			}
		}
		//return false;
	}

	public void handleExcelSeq(File file) {
		try {
			final Workbook workbook = getWorkbook(file);
			final Sheet sheet = workbook.getSheetAt(0);
			handlerSeq = 1;
			String value = null;
/*
			for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
				CellRangeAddress region = sheet.getMergedRegion(i);
				Row firstRow = sheet.getRow(region.getFirstRow());
				Cell firstCellOfFirstRow = firstRow.getCell(region.getFirstColumn());

				if (firstCellOfFirstRow.getCellType() == Cell.CELL_TYPE_STRING) {
					value = firstCellOfFirstRow.getStringCellValue();
				}

				sheet.removeMergedRegion(i);

				for (Row row : sheet) {
					for (Cell cell : row) {
						if (region.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
							cell.setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue(value);
						}
					}
				}

			}*/

			/*for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
				// Delete the region
				final CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
				final int firstColumn = mergedRegion.getFirstColumn();
					//System.out.println("remove merged " + i);
				sheet.removeMergedRegion(i);
			}*/

			while (sheet.getNumMergedRegions() > 0) {
			//	logger.info("Number of merged regions = " + sheet.getNumMergedRegions());
				for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			//		logger.info("Removing merged region " + (i + 1));
					if(sheet.getMergedRegion(i).getFirstColumn() == 0){
						sheet.removeMergedRegion(i);
					}
				}
			}


			final Iterator<Row> rowIterator = sheet.rowIterator();
			int text = 0;
			Row last = null;

			Row startRow = null;
			Row endRow = null;
			for(int i = 0;i <= sheet.getLastRowNum();i++){
				final Row next = sheet.getRow(i);
				if (next == null) {
					if (text == 1) {
						endRow = sheet.getRow(i - 1);
						//do
						doHandleExcelSeq(startRow, endRow);
					}
					text = 0;
					continue;
				}
				final Cell cell1 = next.getCell(0);
				if (cell1 != null) {
					next.removeCell(cell1);
				}
				last = next;
				final Cell cell = next.getCell(1);
				if (cell == null || StringUtils.isBlank(cell.toString().trim())) {
					if (text == 1) {
						endRow = sheet.getRow(next.getRowNum() - 1);
						//do
						doHandleExcelSeq(startRow, endRow);
					}
					text = 0;
				} else {
					if (text == 0) {
						startRow = next;
					}
					text = 1;
				}
			}

			if (text == 1) {
				endRow = last;
				//do
				doHandleExcelSeq(startRow, endRow);
			}


			workbook.write(new FileOutputStream(renamedFile(file)));
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}

	public void handleExcelSeq2(File file) {
		try {
			final Workbook workbook = getWorkbook(file);
			final Sheet sheet = workbook.getSheetAt(0);
			handlerSeq = 1;
			while (sheet.getNumMergedRegions() > 0) {
				for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
					if(sheet.getMergedRegion(i).getFirstColumn() == 0){
						sheet.removeMergedRegion(i);
					}
				}
			}

			int text = 0;
			Row last = null;

			Row startRow = null;
			Row endRow = null;
			for(int i = 0;i <= sheet.getLastRowNum();i++){
				final Row next = sheet.getRow(i);
				last = next;
				final Cell cell = next.getCell(0);
				if (cell != null && cell.toString().matches("\\d+(\\.\\d+)?")) {
					if (text == 0) {
						startRow = next;
						text = 1;
					} else if (text == 1) {
						endRow = sheet.getRow(next.getRowNum() - 1);
						//do
						doHandleExcelSeq(startRow, endRow);
						//text = 0;
						startRow = next;
						endRow = null;
					}
				}
			}

			if (text == 1 && endRow == null) {
				endRow = last;
				//do
				doHandleExcelSeq(startRow, endRow);
			}


			workbook.write(new FileOutputStream(renamedFile(file)));
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}

	public void handleExcelSeq3(File file, String regex) {
		try {
			final Workbook workbook = getWorkbook(file);
			final Sheet sheet = workbook.getSheetAt(0);
			handlerSeq = 1;
			while (sheet.getNumMergedRegions() > 0) {
				for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
					if(sheet.getMergedRegion(i).getFirstColumn() == 0){
						sheet.removeMergedRegion(i);
					}
				}
			}



			int text = 0;
			Row last = null;

			Row startRow = null;
			Row endRow = null;
			for(int i = 0;i <= sheet.getLastRowNum();i++){
				final Row next = sheet.getRow(i);
				last = next;
				final Cell cell1 = next.getCell(0);
				if (cell1 != null) {
					next.removeCell(cell1);
				}
				final Cell cell = next.getCell(1);
				if (cell != null && cell.toString().matches(regex)) {
					if (text == 0) {
						startRow = next;
						text = 1;
					} else if (text == 1) {
						endRow = sheet.getRow(next.getRowNum() - 1);
						//do
						doHandleExcelSeq(startRow, endRow);
						//text = 0;
						startRow = next;
						endRow = null;
					}
				}
			}

			if (text == 1 && endRow == null) {
				endRow = last;
				//do
				doHandleExcelSeq(startRow, endRow);
			}


			workbook.write(new FileOutputStream(renamedFile(file)));
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}

	public String checkExcelBlankRow(File file){
		String text = "没有空行";
		try {
			final Workbook workbook = getWorkbook(file);
			final Sheet sheet = workbook.getSheetAt(0);
			List<Long> nums = new ArrayList<>();
			for(int i = 0;i <= sheet.getLastRowNum();i++){
				final Row row = sheet.getRow(i);
				final Cell cell = row.getCell(0);
				final Cell cell1 = row.getCell(1);
				if (row == null ||  (cell != null && cell1 != null && StringUtils.isBlank(cell.toString()) && StringUtils.isBlank(cell1.toString()))) {
					nums.add((long) (i + 1));
				}
			}
			if (nums.size() > 0) {
				text = "有空行：" + nums;
			}
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
		return text;
	}

	private File renamedFile(File f){
		final String name = f.getName();
		final String parent = f.getParent();
		final int lastIndexOf = name.lastIndexOf('.');
		String newName = name.substring(0, lastIndexOf) + "-" + new SimpleDateFormat("yyyy-MM-dd_HH_mm_ss").format(new Date()) + name.substring(lastIndexOf);
		return new File(parent, newName);
	}


	public static void main(String[] args) {
		//new ExcelHandler().handleExcelSeq(new File("/Users/zyl/Documents/itaojingit/excelhandler/刘斤12-03.xlsx"));
		//new ExcelHandler().handleExcelSeq2(new File("/Users/zyl/Documents/itaojingit/excelhandler/多轮对话12.03(1).xlsx"));
		//new ExcelHandler().handleExcelSeq3(new File("/Users/zyl/Documents/itaojingit/excelhandler/多轮对话12.03(1).xlsx"), "^[^AB]+");
		//final String s = new ExcelHandler().checkExcelBlankRow(new File("/Users/zyl/Documents/itaojingit/excelhandler/多轮对话12.03(1).xlsx"));
		//System.out.println(s);
	}
}
