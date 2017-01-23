package com.sv.excel.test;

import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import com.sv.excel.reader.Excel;

public class Bean extends Excel{
	
	private static Map<String, String> lfMap;
	
	static{
		lfMap = new HashMap<>();
		lfMap.put("名称", "name");
		lfMap.put("编号", "number");
		lfMap.put("日期", "date");
		lfMap.put("大数据", "bdData");
	}
	
	private String name;
	private Integer number;
	private Date date;
	private BigDecimal bdData;
	
	@Override
	public Map<String, String> getLabelField() {
		return lfMap;
	}	

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Integer getNumber() {
		return number;
	}

	public void setNumber(Integer number) {
		this.number = number;
	}

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}

	public BigDecimal getBdData() {
		return bdData;
	}

	public void setBdData(BigDecimal bdData) {
		this.bdData = bdData;
	}

}
