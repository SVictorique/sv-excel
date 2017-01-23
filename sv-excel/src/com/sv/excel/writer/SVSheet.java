package com.sv.excel.writer;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * sv表格页
 * @className SVSheet.java
 * @author 银发Victorique
 * @email 823245670@qq.com
 * @date 2017年1月23日
 */
public class SVSheet {

	private Workbook wb;//POI工作簿对象
	private Sheet sheet;//POI表格页对象
	private Row row;//POI行对象
	
	private int columns;//表格页总列数，用于自动换行
	
	private int rownum = 0;//行游标，用于记录当前插入单元格的行数
	private int colnum = 0;//列游标，用于记录当前插入单元格的列数
	private List<String> mergeBuffer = new ArrayList<>();//用于保存进行过合并单元格的单元格坐标，数据格式：行-列
	
	/**
	 * 构造函数，用于初始化工作簿、表格页、总列数
	 * @param wb POI工作簿对象
	 * @param sheet POI表格页对象
	 * @param columns 表格页最大列数
	 */
	protected SVSheet(Workbook wb, Sheet sheet, int columns){
		this.wb = wb;
		this.sheet = sheet;
		this.columns = columns;
	}
	
	/**
	 * 添加sv单元格
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @param svCell
	 */
	public void addCell(SVCell svCell){
		//判断单元格是否被合并，并调整游标位置
		String key = String.valueOf(this.rownum)+"-"+String.valueOf(this.colnum);
		while(mergeBuffer.contains(key)){
			colnum ++;
			key = String.valueOf(this.rownum)+"-"+String.valueOf(this.colnum);
		}
		//判断是否需要创建行
		if(row == null || colnum >= columns){
			row = this.sheet.createRow(rownum);
			rownum ++;
			colnum = 0;
		}
		if(svCell.getRowspan() > 1 || svCell.getColspan() >1){
			//合并单元格
			if(row != null){
				sheet.addMergedRegion(new CellRangeAddress(rownum-1, rownum-1+svCell.getRowspan()-1, colnum, colnum+svCell.getColspan()-1));
			}else{
				sheet.addMergedRegion(new CellRangeAddress(rownum, rownum+svCell.getRowspan()-1, colnum, colnum+svCell.getColspan()-1));
			}
			//保存合并的单元格坐标
			for(int i=0;i<svCell.getRowspan();i++){
				for(int j=0;j<svCell.getColspan();j++){
					String value = String.valueOf(this.rownum+i)+"-"+String.valueOf(this.colnum+j);
					mergeBuffer.add(value);
				}
			}
		}
		//创建单元格
		Cell cell = row.createCell(colnum);
		CellStyle cellStyle = null;
		//创建单元格样式
		if(svCell.getCellStyle() == null){
			cellStyle = wb.createCellStyle();;
		}else{
			cellStyle = svCell.getCellStyle();
		}
		//获取单元格的值
		Object value = svCell.getCellValue();
		//单元格赋值
		if(value instanceof Double){
			cell.setCellValue((Double)value);
		}else if(value instanceof Date){
			CreationHelper createHelper = wb.getCreationHelper();
		    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
			cell.setCellValue((Date)value);
		}else if(value instanceof Calendar){
			CreationHelper createHelper = wb.getCreationHelper();
		    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
			cell.setCellValue((Calendar)value);
		}else if(value instanceof RichTextString){
			cell.setCellValue((RichTextString)value);
		}else if(value instanceof String){
			cell.setCellValue((String)value);
		}
		//添加单元格样式
		cell.setCellStyle(cellStyle);
		colnum ++;
		//设置MyCell对象
		svCell.setCell(cell);
	}
	
	/**
	 * 获取POI表格页对象，此方法用于直接获取POI表格页对象，使用POI原生的方法进行操作
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return
	 */
	public Sheet getSheet(){
		return this.sheet;
	}
	
}
