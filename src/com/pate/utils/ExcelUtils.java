package com.pate.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import com.pate.excel.CellData;

public class ExcelUtils {

	public static HSSFWorkbook openWorkbook(String filePath) throws IOException {
		File file = null;
		file = new File(filePath);
		FileInputStream ins = null;
		HSSFWorkbook workbook;
		ins = new FileInputStream(file);
		workbook = new HSSFWorkbook(ins);
		return workbook;
	}
	
	public static String getCellString(Cell cell) {
		if (cell == null) return null;
		int type = cell.getCellType();
		if (type == Cell.CELL_TYPE_BLANK) {
			return "";
		} else if (type == Cell.CELL_TYPE_BOOLEAN) {
			return cell.getBooleanCellValue() ? "true" : "false";
		} else if (type == Cell.CELL_TYPE_ERROR) {
			return Byte.toString(cell.getErrorCellValue());
		} else if (type == Cell.CELL_TYPE_FORMULA) {
			double n = cell.getNumericCellValue();
			return java.math.BigDecimal.valueOf(n).toPlainString();
		} else if (type == Cell.CELL_TYPE_NUMERIC) {
			double n = cell.getNumericCellValue();
			return java.math.BigDecimal.valueOf(n).toPlainString();
		} else if (type == Cell.CELL_TYPE_STRING) {
			return cell.getStringCellValue();
		} else {
			return null;
		}
	}
	
	public static double getCellDouble(Cell cell) throws Exception {
		int type = cell.getCellType();
		if (type == Cell.CELL_TYPE_BLANK) {
			return 0;
		} else if (type == Cell.CELL_TYPE_BOOLEAN) {
			return cell.getBooleanCellValue() ? 1 : 0;
		} else if (type == Cell.CELL_TYPE_ERROR) {
			return Byte.valueOf(cell.getErrorCellValue()).doubleValue();
		} else if (type == Cell.CELL_TYPE_FORMULA) {
			return cell.getNumericCellValue();
		} else if (type == Cell.CELL_TYPE_NUMERIC) {
			return cell.getNumericCellValue();
		} else if (type == Cell.CELL_TYPE_STRING) {
			return 0;
		} else {
			return 0;
		}
	}
	
	public static CellRangeAddress getCellRageAddress(Cell cell) {
		Sheet sheet = cell.getSheet();
		int num = sheet.getNumMergedRegions();
		if (num < 1) return null;
		
		CellRangeAddress addr = null;
		for (int i = 0; i < num; i++) {
			addr = sheet.getMergedRegion(i);
			if (addr.getFirstRow() == cell.getRowIndex() &&
					addr.getFirstColumn() == cell.getColumnIndex()) {
				return addr;
			}
		}
		
		return null;
	}
	
	public static short getCellAlignment(Cell cell) {
		CellStyle style = cell.getCellStyle();
		return style.getAlignment();
	}
	
	public static short getCellVerticalAlignment(Cell cell) {
		CellStyle style = cell.getCellStyle();
		return style.getVerticalAlignment();
	}
	
	public static CellData getCellData(Cell cell) throws Exception {
		CellData data = null;
		if (cell == null) return data;
		
		data = new CellData();
		data.setCellIndex(cell.getColumnIndex());
		data.setRowIndex(cell.getRowIndex());
		int type = cell.getCellType();
		switch (type) {
		case Cell.CELL_TYPE_BLANK:
			data.setDataType(CellData.STRING_CELL);
			data.setStringValue("");
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			data.setDataType(CellData.NUMBERIC_CELL);
			data.setNumberValue(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			data.setDataType(CellData.NUMBERIC_CELL);
			data.setNumberValue(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			data.setDataType(CellData.NUMBERIC_CELL);
			data.setNumberValue(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			data.setDataType(CellData.NUMBERIC_CELL);
			data.setNumberValue(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			data.setDataType(CellData.STRING_CELL);
			data.setStringValue(cell.getStringCellValue());
			break;
		}
		
		return data;
	}

	public static void printCellDataList(List<CellData> list) {
		for (CellData e : list) {
			if (e.getDataType() == CellData.STRING_CELL) {
				System.out.println("row:" + e.getRowIndex() +
						" column:" + e.getCellIndex() +
						" type:" + e.getDataType() +
						" value:" + e.getStringValue());
			} else {
				System.out.println("row:" + e.getRowIndex() +
						" column:" + e.getCellIndex() +
						" type:" + e.getDataType() +
						" value:" +
						BigDecimal.valueOf(e.getNumberValue()).toPlainString());
				
			}
		}
	}
	
	public static String getCellString(HSSFSheet sheet, String pos) {
		int cellIndex = pos.charAt(0) - 'A';
		int rowIndex = Integer.valueOf(pos.substring(1));
		Row row = sheet.getRow(rowIndex);
		if (row == null) return "";
		Cell cell = row.getCell(cellIndex);
		if (cell == null) return "";
		return getCellString(cell);
	}
	
	public static CellData getCellData(HSSFSheet sheet, int rowIndex, int colIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) return null;
		return getCellData(row, rowIndex, colIndex);
	}
	public static CellData getCellData(Row row, int rowIndex, int colIndex) {
		Cell cell = row.getCell(colIndex);
		if (cell == null) return null;
		int cellType = cell.getCellType();
		CellData data = new CellData();
		data.setRowIndex(rowIndex);
		data.setCellIndex(colIndex);
		data.setDataType(cellType);
		switch (cellType) {
		case Cell.CELL_TYPE_STRING :
			data.setValue(cell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_BOOLEAN :
			boolean bval = cell.getBooleanCellValue();
			data.setValue(Boolean.valueOf(bval));
			break;
		case Cell.CELL_TYPE_FORMULA :
		case Cell.CELL_TYPE_NUMERIC :
			double dval = cell.getNumericCellValue();
			data.setValue(Double.valueOf(dval));
			break;
		case Cell.CELL_TYPE_ERROR :
			byte btVal = cell.getErrorCellValue();
			data.setValue(Byte.valueOf(btVal));
			break;
		default :
			break;	
		}
		return data;
	}
	
	
}
