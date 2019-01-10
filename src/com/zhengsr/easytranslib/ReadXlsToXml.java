package com.zhengsr.easytranslib;

import java.awt.Frame;
import java.awt.LayoutManager;
import java.awt.List;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;

import com.zhengsr.easytranslib.bean.XlsReadBean;
import com.zhengsr.easytranslib.method.ReadXlsManager;



/**
 * @author zhengshaorui 2018/6/24
 */
public class ReadXlsToXml {

	private static String ROOT_FILE = "test.xlsx"; // excel 的名字
	private static String FILE_NAME = "Demo";  //生成的文件夹的子
	public static String STRING_NAME = "test_strings.xml"; //要生成的 strings 的名字
	public static String ARRAY_NAME = "test_arrays.xml"; //要生成的 array 的名字
	private static String ROOT_PATH; // 当前路径
	private static int IGNORE_ROW = 1; //要忽略的行,根据你们的 xls 来

	public static void main(String[] args) {
		File file = new File("");
		ROOT_PATH = file.getAbsolutePath();
		XlsReadBean bean = new XlsReadBean.Builder()
			.setRootPath(ROOT_PATH) //根路径
			.setXlsFile(ROOT_FILE) //xls文件
			.setFileFloderName(FILE_NAME) //要生成的文件夹名字
			.setIgnoreRow(IGNORE_ROW) //需要忽略的行，即不需要转换的
			.setStringName(STRING_NAME) //strings.xml 的名字，可以客制化
			.setArrayName(ARRAY_NAME) //array.xml 的名字，可以客制化
			.builder();
	
		ReadXlsManager.getInstance().readXls(bean.getBuilder());
		System.out.println("在 "+ROOT_PATH+File.separator+FILE_NAME+" 生成文件啦!!");
	}
	
	


	

	
}
