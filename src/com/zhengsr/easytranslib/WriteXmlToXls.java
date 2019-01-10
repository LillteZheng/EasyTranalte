package com.zhengsr.easytranslib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import com.zhengsr.easytranslib.bean.XlsWriteBean;
import com.zhengsr.easytranslib.method.WriteXlsManager;

/**
 * @author zhengshaorui 2018/6/24
 */
public class WriteXmlToXls {
	private static String VALUE_PATH = "Test";  //value文件夹存放的名字
	private static String XLS_NAME = "test.xlsx"; //要生成xls的名字
	private static String ROOT_PATH; // 当前路径

	
	public static void main(String[] args) {
		File file = new File("");
		ROOT_PATH = file.getAbsolutePath();
		XlsWriteBean bean = new XlsWriteBean.Builder()
				.setRootPath(ROOT_PATH)
				.setFileFloderName(VALUE_PATH)
				.setXlsName(XLS_NAME)
				.builder();
		WriteXlsManager.getInstance().startWrite(bean.getBuilder());
		System.out.println("在 "+ROOT_PATH+File.separator+VALUE_PATH+"生成文件啦!!");
	}


}	
