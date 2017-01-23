package com.sv.excel.reader;

import java.util.Map;

/**
 * Excel抽象类，用于读取数据为对象集合
 * @className Excel.java
 * @author 银发Victorique
 * @email 823245670@qq.com
 * @date 2017年1月23日
 */
public abstract class Excel {

	/**
	 * 获取存放表头-字段对应关系表
	 * 数据格式为：{表头名称:字段名称}
	 * 阅读器将根据该关系表给字段赋值
	 * @author 银发Victorique
	 * @email 823245670@qq.com
	 * @return 表头-字段对应关系表
	 */
	public abstract Map<String, String> getLabelField();
	
}
