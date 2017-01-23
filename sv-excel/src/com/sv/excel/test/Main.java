package com.sv.excel.test;

import java.util.Calendar;
import java.util.List;

import com.sv.excel.reader.SVExcelReader;
import com.sv.excel.writer.SVCell;
import com.sv.excel.writer.SVSheet;
import com.sv.excel.writer.SVWorkbook;

public class Main {

	public static void main(String[] args) {
		/*SVWorkbook wb = new SVWorkbook("./", "workbook.xlsx");
		SVSheet sheet = wb.createSheet(4);
		
		SVCell cell = new SVCell();
		cell.setCellValue("test");
		sheet.addCell(cell);
		
		cell = new SVCell();
		cell.setCellValue(Calendar.getInstance());
		cell.setColspan(3);
		cell.setRowspan(2);
		sheet.addCell(cell);
		
		cell = new SVCell();
		cell.setCellValue(123.231);
		sheet.addCell(cell);
		
		wb.close();*/
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
		/*List<Bean> rows = SVExcelReader.getBeanList("./", "workbook.xlsx", 0, new Bean());
		for (Bean bean : rows) {
			System.out.println(bean);
		}*/
	}

}
