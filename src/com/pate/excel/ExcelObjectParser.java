package com.pate.excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;

import com.pate.utils.ExcelUtils;


public class ExcelObjectParser {
	
	public static List<Object> parse(HSSFSheet sheet, ObjectRule rule) {
		List<ObjectWrapper> objectList = new ArrayList<ObjectWrapper>();
		
		List<FieldRule> fieldRules = rule.getFieldRuleList();
		for (FieldRule fieldRule : fieldRules) {
			readField(sheet, rule.getObjectClass(), fieldRule, objectList);
		}
		
		if (objectList.size() > 0 && objectList.get(0).getPosition() == FieldRule.NONE) {
			Object obj = objectList.get(0).getObject();
			for (FieldRule fieldRule : fieldRules) {
				if (fieldRule.getPositionType() != FieldRule.NONE) continue;
				Field field = fieldRule.getField();
				field.setAccessible(true);
				try {
					Object fieldVal = field.get(obj);
					if (fieldVal != null) {
						for (ObjectWrapper ow : objectList) {
							field.set(ow.getObject(), fieldVal);
						}
					}
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				field.setAccessible(false);
			}
		}

		List<Object> list = new ArrayList<Object>();
		for (ObjectWrapper ow : objectList) {
			if (ow.getPosition() == FieldRule.NONE) continue;
			list.add(ow.getObject());
		}
		return list;
	}
	
	private static <T> void readField(
			HSSFSheet sheet,
			Class<T> objectClass,
			FieldRule fieldRule,
			List<ObjectWrapper> objectList) {
		int startRow = fieldRule.getStartRow();
		int endRow = fieldRule.getEndRow();
		int startCol = fieldRule.getStartCol();
		int endCol = fieldRule.getEndCol();
		int positionType = fieldRule.getPositionType();
		int position;
		ObjectWrapper objectWrapper = null;

		for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			for (int col = startCol; col <= endCol; col++) {
				CellData cellData = ExcelUtils.getCellData(row, rowIndex, col);
				if (cellData == null || cellData.getValue() == null) continue;
				position = getPosition(cellData, positionType);
				objectWrapper = getObjectWrapper(objectList, objectClass, position);
				setValue(objectWrapper.getObject(), fieldRule.getField(),
						cellData.getValue());
			}
		}
	}
	
	private static int getPosition(CellData cellData, int positionType) {
		if (positionType == FieldRule.COLUMN_POSITION) {
			return cellData.getCellIndex();
		} else if (positionType == FieldRule.ROW_POSITION) {
			return cellData.getRowIndex();
		} else {
			return FieldRule.NONE;
		}
	}
	
	private static <T> ObjectWrapper newObjectWrapper(Class<T> objectClass, int position) {
		ObjectWrapper ow = new ObjectWrapper();
		Object obj = null;
		try {
			obj = objectClass.newInstance();
		} catch (InstantiationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		ow.setObject(obj);
		ow.setPosition(position);
		return ow;
	}
	
	private static void setValue(Object object, Field field,
			Object fieldVal) {
			try {
				field.set(object, fieldVal);
			} catch (IllegalArgumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	}
	
	private static <T> ObjectWrapper getObjectWrapper(List<ObjectWrapper> list,
			Class<T> objectClass, int position) {
		if (position == -1) {
			if (list.size() == 0 || list.get(0).getPosition() != -1) {
				
				ObjectWrapper ow = newObjectWrapper(objectClass, position);
				list.add(0, ow);
			}
			return list.get(0);
		}
		
		for (ObjectWrapper objectWrapper : list) {
			if (position == objectWrapper.getPosition()) {
				return objectWrapper;
			}
		}
		
		ObjectWrapper ow = newObjectWrapper(objectClass, position);
		list.add(ow);
		return ow;
	}
}
