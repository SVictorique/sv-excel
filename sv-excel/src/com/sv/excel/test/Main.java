package com.sv.excel.test;

import java.util.List;
import java.util.Map;

import com.sv.excel.reader.SVExcelReader;

public class Main {

	public static void main(String[] args) {
		//getOriginalArray方法测试
		/*Object[][] rows = SVExcelReader.getOriginalArray("./", "workbook.xlsx", 0);
		for (Object[] cols : rows) {
			for (Object obj : cols) {
				System.out.print(obj+"\t");
			}
			System.out.println();
		}*/

		//getMapList方法测试
		/*List<Map<String, Object>> rows = SVExcelReader.getMapList("./", "workbook.xlsx", 0);
		for (Map<String, Object> cols : rows) {
			System.out.println(cols);
		}*/

		//getBeanList方法测试
		List<Bean> rows = SVExcelReader.getBeanList("./", "workbook.xlsx", 0, new Bean());
		for (Bean bean : rows) {
			System.out.println(bean);
		}
	}

}
