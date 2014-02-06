package com.pate.excel;

import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.*;

import com.pate.utils.ExcelUtils;

public class Test {
	
	private static void test01() {
		HSSFWorkbook wb = null;
		try {
			wb = ExcelUtils.openWorkbook(".\\doc\\�۸��.xls");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return;
		}
		List<ObjectRule> ruleList = null;
		try {
			ruleList = Rules.getRules();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if (ruleList == null) return;
		HSSFSheet sheet = wb.getSheet("ȡ��");
		if (sheet != null) {
			for (ObjectRule rule : ruleList) {
				List list = ExcelObjectParser.parse(sheet, rule);
				Test.printPriceTable(list);
			}
		}
	}

	public static void main(String[] args) {
		test01();
	}
	    
	private static void printPriceTable(List<PriceTable> list) {
		for (int i = 0; i < list.size(); i++) {
			PriceTable c = list.get(i);
			// �۸�Ϊ��Ĳ����
			if (c.getPrice() - 0 < 0.00000001) continue;
			System.out.println("���ţ�" + c.getBuildingNumber() +
					" ¥�㣺" + c.getFloor() +
					" ���ţ�" + c.getRoomNumber() +
					" �۸�" + Double.valueOf(c.getPrice()).toString());
		}
	}
	
}