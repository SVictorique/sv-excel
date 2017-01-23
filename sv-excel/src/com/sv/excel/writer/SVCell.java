package com.sv.excel.writer;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 * sv单元格
 * @className SVCell.java
 * @author 银发Victorique
 * @email 823245670@qq.com
 * @date 2017年1月23日
 */
public class SVCell {

	private Cell cell;//POI单元格对象
	
	private int colspan = 1;//跨列
	private int rowspan = 1;//跨行
	
	private Double doubleValue;//双精度浮点数类型数据
	private Date dateValue;//日期类型数据
	private Calendar calendarValue;//日历类型数据
	private RichTextString richTextStringValue;//富文本类型数据
	private String stringValue;//字符串类型数据
	
	private CellStyle cellStyle;//POI单元格样式
	
	/**
	 * 设置双精度浮点数类型的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param doubleValue
	 * @return sv单元格对象
	 */
	public SVCell setCellValue(double doubleValue){
		this.resetValues();
		this.doubleValue = doubleValue;
		return this;
	}

	/**
	 * 设置日期类型的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param dateValue
	 * @return sv单元格对象
	 */
	public SVCell setCellValue(Date dateValue){
		this.resetValues();
		this.dateValue = dateValue;
		return this;
	}

	/**
	 * 设置日历格式的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param calendarValue
	 * @return sv单元格对象
	 */
	public SVCell setCellValue(Calendar calendarValue){
		this.resetValues();
		this.calendarValue = calendarValue;
		return this;
	}

	/**
	 * 设置POI富文本类型的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param richTextStringValue
	 * @return sv单元格对象
	 */
	public SVCell setCellValue(RichTextString richTextStringValue){
		this.resetValues();
		this.richTextStringValue = richTextStringValue;
		return this;
	}

	/**
	 * 设置字符串类型的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param stringValue
	 * @return sv单元格对象
	 */
	public SVCell setCellValue(String stringValue){
		this.resetValues();
		this.stringValue = stringValue;
		return this;
	}
	
	/**
	 * 获取单元格的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return 返回不为null的值
	 */
	public Object getCellValue(){
        if(doubleValue != null){
        	return this.doubleValue;
        }else if(dateValue != null){
        	return this.dateValue;
        }else if(calendarValue != null){
        	return this.calendarValue;
        }else if(richTextStringValue != null){
        	return this.richTextStringValue;
        }else if(stringValue != null){
        	return this.stringValue;
        }else{
        	return null;
        }
	}
	
	/**
	 * 设置单元格样式
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param cellStyle POI单元格样式
	 * @return sv单元格
	 */
	public SVCell setCellStyle(CellStyle cellStyle){
		this.cellStyle = cellStyle;
		return this;
	}
	
	/**
	 * 获取单元格样式
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return POI单元格样式
	 */
	public CellStyle getCellStyle(){
		return this.cellStyle;
	}
	
	/**
	 * 获取POI单元格对象，此方法用于直接获取POI单元格对象，使用POI原生的方法进行操作
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return POI单元格对象
	 */
	public Cell getCell(){
		return this.cell;
	}
	
	/**
	 * 设置跨列
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param colspan 跨列数量，默认为1
	 * @return sv单元格
	 */
	public SVCell setColspan(int colspan){
		this.colspan = colspan;
		return this;
	}
	
	/**
	 * 设置跨行
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param rowspan 跨行数量，默认为1
	 * @return sv单元格
	 */
	public SVCell setRowspan(int rowspan){
		this.rowspan = rowspan;
		return this;
	}
	
	/**
	 * 设置单元格
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param cell POI单元格
	 */
	protected void setCell(Cell cell){
		this.cell = cell;
	}
	
	/**
	 * 获取跨列
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return 跨列数
	 */
	protected int getColspan(){
		return this.colspan;
	}
	
	/**
	 * 获取跨行
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return 跨行数
	 */
	protected int getRowspan(){
		return this.rowspan;
	}
	
	/**
	 * 重置单元格的值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 */
	private void resetValues(){
		doubleValue = null;
		dateValue = null;
		calendarValue = null;
		richTextStringValue = null;
		stringValue = null;
	}
	
}
