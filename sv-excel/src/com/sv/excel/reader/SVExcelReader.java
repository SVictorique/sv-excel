package com.sv.excel.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * svExcel阅读器
 * @className SVExcelReader.java
 * @author 银发Victorique
 * @email 823245670@qq.com
 * @date 2017年1月23日
 */
public class SVExcelReader {

	/**
	 * 获取原始数据类型的数组
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param path 文件路径
	 * @param filename 文件名称
	 * @param sheetnum 表格页编号，从0开始
	 * @return 表格数据数组
	 */
	public static Object[][] getOriginalArray(String path, String filename, int sheetnum){
		Object[][] returnArray = null;
		InputStream is = null;
		Workbook wb = null;
		try {
			is = new FileInputStream(path+"/"+filename);
			if(filename.endsWith(".xls")){
				wb = new HSSFWorkbook(is);
			}else if(filename.endsWith(".xlsx")){
				wb = new XSSFWorkbook(is);
			}else{
				return null;
			}
			int maxcol = 0;
			for(Row row : wb.getSheetAt(sheetnum)){
				if(row.getPhysicalNumberOfCells() > maxcol){
					maxcol = row.getPhysicalNumberOfCells();
				}
			}
			returnArray = new Object[wb.getSheetAt(sheetnum).getPhysicalNumberOfRows()][maxcol];
			for(int i=0; i<wb.getSheetAt(sheetnum).getPhysicalNumberOfRows(); i++){
				Row row = wb.getSheetAt(sheetnum).getRow(i);
				for(int j=0; j<maxcol; j++){
					Cell cell = row.getCell(j);
					if(cell != null){
						returnArray[i][j] = getCellValue(cell);
					}else{
						returnArray[i][j] = "";
					}
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if(wb != null){
					wb.close();
				}
			} catch (IOException e1) {
				e1.printStackTrace();
			} finally {
				if(is != null){
					try {
						is.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
		return returnArray;
	}
	
	/**
	 * 获取List<Map<String,Object>>类型的表格数据
	 * 要求导入的EXCEL文件有这样的格式：
	 * 第一行为表头，必须全部为字符串型数据，否则会抛出异常
	 * 之后的每一行的数据，都会以其对应的表头作为键存储在map中
	 * 如果某一单元格对应的表头为空，则此单元格数据被舍弃
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param path 文件路径
	 * @param filename 文件名称
	 * @param sheetnum 表格页编号，从0开始
	 * @return 表格页数据
	 */
	public static List<Map<String, Object>> getMapList(String path, String filename, int sheetnum){
		List<Map<String, Object>> returnList = new ArrayList<>();
		InputStream is = null;
		Workbook wb = null;
		try {
			is = new FileInputStream(path+"/"+filename);
			if(filename.endsWith(".xls")){
				wb = new HSSFWorkbook(is);
			}else if(filename.endsWith(".xlsx")){
				wb = new XSSFWorkbook(is);
			}else{
				return null;
			}
			Row firstRow = wb.getSheetAt(sheetnum).getRow(0);
			for(int i=1; i<wb.getSheetAt(sheetnum).getPhysicalNumberOfRows(); i++){
				Row row = wb.getSheetAt(sheetnum).getRow(i);
				Map<String, Object> map = new HashMap<>();
				for(int j=0; j<row.getPhysicalNumberOfCells(); j++){
					Cell cell = row.getCell(j);
					map.put(firstRow.getCell(j).getStringCellValue(), getCellValue(cell));
				}
				returnList.add(map);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if(wb != null){
					wb.close();
				}
			} catch (IOException e1) {
				e1.printStackTrace();
			} finally {
				if(is != null){
					try {
						is.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
		return returnList;
	}
	
	/**
	 * 获取bean集合类型的表格页数据
	 * 要求导入的EXCEL文件有这样的格式：
	 * 第一行为表头，必须全部为字符串型数据，否则会抛出异常
	 * 实例excel必须是抽象类Excel的子类的实例，并且getLabelField方法获取的Map需要与表头相对应，否则会被舍弃
	 * 之后的每一行的数据，都会以其对应的表头和getLabelField获取的Map作为依据，存放到Bean的对应字段中
	 * 如果某一单元格对应的表头为空，则此单元格数据被舍弃
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param path 文件路径
	 * @param filename 文件名
	 * @param sheetnum 表格页编号，从0开始
	 * @param excel 抽象类Excel的子类的实例
	 * @return 表格页数据
	 */
	public static <T extends Excel> List<T> getBeanList(String path, String filename, int sheetnum, T excel){
		List<T> returnList = new ArrayList<>();
		Map<String, String> lfMap = excel.getLabelField();
		InputStream is = null;
		Workbook wb = null;
		try {
			is = new FileInputStream(path+"/"+filename);
			if(filename.endsWith(".xls")){
				wb = new HSSFWorkbook(is);
			}else if(filename.endsWith(".xlsx")){
				wb = new XSSFWorkbook(is);
			}else{
				return null;
			}
			Row firstRow = wb.getSheetAt(sheetnum).getRow(0);
			for(int i=1; i<wb.getSheetAt(sheetnum).getPhysicalNumberOfRows(); i++){
				Row row = wb.getSheetAt(sheetnum).getRow(i);
				@SuppressWarnings("unchecked")
				T instance = (T) excel.getClass().getConstructor().newInstance();
				for(int j=0; j<row.getPhysicalNumberOfCells(); j++){
					Cell cell = row.getCell(j);
					String field = lfMap.get(firstRow.getCell(j).getStringCellValue());
					if(field != null){
						Object value = getCellValue(cell);
						Class<?> clazz = null;{
							if(value instanceof RichTextString){
								clazz = String.class;
								value = ((RichTextString)value).getString();
							}else{
								clazz = value.getClass();
							}
						}
						try {
							Method method = instance.getClass().getMethod("set"+field.substring(0,1).toUpperCase()+field.substring(1), clazz);
							method.invoke(instance, value);
						} catch (Exception e) {
							if(value instanceof Double){
								try {
									Method method = instance.getClass().getMethod("set"+field.substring(0,1).toUpperCase()+field.substring(1), Integer.class);
									method.invoke(instance, ((Double)value).intValue());
								} catch(Exception e1){
									Method method = instance.getClass().getMethod("set"+field.substring(0,1).toUpperCase()+field.substring(1), BigDecimal.class);
									BigDecimal bd = new BigDecimal((Double)value);
									method.invoke(instance, bd);
								}
							}else{
								e.printStackTrace();
							}
						}
					}
				}
				returnList.add(instance);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		} finally {
			try {
				if(wb != null){
					wb.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				if(is != null){
					try {
						is.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
		return returnList;
	}

	/**
	 * 获取单元格原始数据
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param cell 单元格
	 * @return 原始数据
	 */
	@SuppressWarnings("deprecation")
	private static Object getCellValue(Cell cell){
		if(cell == null){
			return null;
		}
		switch (cell.getCellTypeEnum()) {
		case STRING:
			return cell.getRichStringCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case FORMULA:
			return cell.getCellFormula();
		case BLANK:
			return null;
		default:
			return null;
		}
	}

}
