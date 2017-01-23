package com.sv.excel.writer;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * sv工作簿
 * @className SVWorkbook.java
 * @author 银发Victorique
 * @email 823245670@qq.com
 * @date 2017年1月23日
 */
public class SVWorkbook {
	
	private Workbook workbook;//POI工作簿对象
	
	private FileOutputStream fos;//文件输出流
	
	/**
	 * 构造函数，根据路径和文件名创建工作簿和输出流
	 * @param path 文件路径（不包含文件名称），例如：D:\sv\excel\demo\
	 * @param fileName 文件名，例如：svexceldemo.xls/svexceldemo.xlsx
	 */
	public SVWorkbook(String path, String fileName){
		//判断文件类型，创建工作簿对象
		if(fileName.endsWith(".xls")){
			workbook = new HSSFWorkbook();
		}else if(fileName.endsWith(".xlsx")){
			workbook = new XSSFWorkbook();
		}else{
			return;
		}
		//根据文件路径和文件名称，创建文件输出流
		try {
			fos = new FileOutputStream(path+"/"+fileName);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 创建sv表格页，为了更方便的进行行和单元格的添加，创建表格页时，需要填写最大列数，操作时只需添加单元格，换行等操作自动完成
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param columns 表格的最大列数
	 * @return sv表格页对象
	 */
	public SVSheet createSheet(int columns){
		return new SVSheet(this.workbook, this.workbook.createSheet(), columns);
	}
	
	/**
	 * 关闭，将创建的工作簿内容写入文件，同时关闭文件输出流
	 * 注：完成创建全部内容后，必须执行此方法，否则会造成数据没有写入，流没有关闭等严重问题
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 */
	public void close(){
		if(workbook != null){
			//写入内容
			try {
				workbook.write(fos);
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				if(fos != null){
					//关闭文件流
					try {
						fos.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
	}
	
	/**
	 * 获取POI工作簿对象，此方法用于直接获取POI工作簿对象，使用POI原生的方法进行操作
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return POI工作簿对象
	 */
	public Workbook getWorkBook(){
		return this.workbook;
	}
	
}
