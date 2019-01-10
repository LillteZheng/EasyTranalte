package com.zhengsr.easytranslib.bean;

import java.util.List;

/**
 * @author zhengshaorui 2018/6/24
 */
public class CusRow {
	public String key;
	public String value;
	public boolean isArray;
	public List<String> items;
	@Override
	public String toString() {
		return "CusRow [key=" + key + ", value=" + value + ", isArray="
				+ isArray + ", items=" + items + "]";
	}
	
	
}
