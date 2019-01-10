package com.zhengsr.easytranslib.method;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import com.zhengsr.easytranslib.bean.XlsReadBean;




/**
 * @author zhengshaorui 2018/6/24
 */
public class ReadXlsManager extends BaseSheet{
	private static int MAX_COLUME;
	private static int MAX_ROW;
	private XlsReadBean.Builder mBuilder;
	private boolean isArrayFile = false;
	private Map<Integer, String> mMap = new HashMap<>();
	private String mStringKey;
	private boolean isLastItemString = false;
	private static class Holder {
		static ReadXlsManager INSTANCE = new ReadXlsManager();
	}

	public static ReadXlsManager getInstance() {
		return Holder.INSTANCE;
	}
	
	private ReadXlsManager(){
		
	}
	
	public void readXls(XlsReadBean.Builder builder){
		mBuilder = builder;
		File file = new File(builder.getRootPath(),builder.getXlsFile());
		if(file.exists()){
			try {
				InputStream is = new FileInputStream(file);
				Workbook wb = null;
				//判断 xls 的版本
				if(builder.getXlsFile().toLowerCase().endsWith("xlsx")){
					wb = new XSSFWorkbook(is);
				}else{
					wb = new HSSFWorkbook(is);
				}
				int sheetNum = wb.getNumberOfSheets();
				for (int i = 0; i < sheetNum; i++) { //看有多少个 sheet
					Sheet sheet = (Sheet) wb.getSheetAt(i);
					int firstRowIndex = sheet.getFirstRowNum(); // 第一行
					int lastRowIndex = sheet.getLastRowNum(); // 最后一行
					
					MAX_ROW = getMaxRow(sheet)+1;
					
					
					for (int j = firstRowIndex; j < MAX_ROW; j++) { //行
						Row row = sheet.getRow(j);
						if (j == 0) {
							readFirstRow(row);
						} else {
							readAllRows(row, j);
						}
					}
				}
			} catch (Exception e) {
				// TODO: handle exception
				System.out.println("error: "+e.toString());
			}
		}else{
			System.out.println("Cannot find "+builder.getRootPath()+File.separator+builder.getXlsFile());
		}
	}
	
	/**
	 * 获取xml最大行数
	 * 
	 * @param sheet
	 * @return
	 */
	public int getMaxRow(Sheet sheet) {
		int firstrow = sheet.getFirstRowNum();
		int lastrow = sheet.getLastRowNum();
		int num = 0;
		for (int i = firstrow; i < lastrow; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				num++;
			}
		}
		return num;
	}
	
	/**
	 * 读取第一行，并创建对应的文件夹
	 * @param row
	 * @param maxRow
	 */
	private void readFirstRow(Row row){
		if (row != null) {
			int firstCellIndex = row.getFirstCellNum();
			MAX_COLUME = row.getLastCellNum();
			
			// 此处参数cIndex决定可以取到excel的列数。
			for (int j = firstCellIndex ; j < MAX_COLUME; j++) { //列
				Cell cell = row.getCell(j);
				String value = "";
				if(j == 0){ //第一行，第一列，判断是 string 类型还是 array类型
					if(cell != null){
						value = cell.toString();
						if(value.contains("type_array")){
							isArrayFile = true;
						}
					}
				}else{
					if (cell != null) {
						value = cell.toString();
						//不同语言对应的文件夹名称
						String flodername = getFloderByLang(value);
						System.out.println("语言: "+value+" 文件夹: "+flodername);
						
						
						//创建路径
						String filepath = createFloder(mBuilder.getRootPath(), mBuilder.getFileFloderName()+ "/" + flodername);
						mMap.put(j,filepath);
						//创建文件
						createFileAndData(mMap.get(j),isArrayFile,mBuilder);
					}
				}
			}	
		}
	}
	
	/**
	 * 读取每一行的数据，并写到对应的文件夹中
	 * @param row
	 * @param rowIndex
	 */
	private  void readAllRows(Row row,int rowIndex){
		
		if (row != null) {
			int firstCellIndex = row.getFirstCellNum();
			int lastCellIndex = row.getLastCellNum();
			// 此处参数cIndex决定可以取到excel的列数。
			for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) {
				Cell cell = row.getCell(cIndex);
				if (cell != null) {
					if(isArrayFile){
						writeValueToArray(cIndex, cell, rowIndex);
					}else{
						String path = mMap.get(cIndex);
						writeValueToString(cIndex, cell, rowIndex, path);
					}

				}
			}
		}
	}
	
	
	/**
	 * 写数据到 string 文件上
	 * @param cIndex 列
	 * @param cell 单元格
	 * @param rowIndex 行
	 * @param dir 路径
	 */
	public void writeValueToString(int cIndex,Cell cell, int rowIndex,String dir) {
		String value = "";
		if(cell != null){
			value = cell.toString();
			if(cIndex == 0){ //先保留key值
				mStringKey = value;
				//System.out.println("key: "+mStringKey+" "+mStringKey.length()+" "+mStringKey.isEmpty());
				/*if(mStringKey.isEmpty()){
					mStringKey = "display_sleep_mode_summary";
				}*/
			}else{ //然后把数据写进去
				int type = cell.getCellType();
				//日期需要特殊处理
				/*if(type == HSSFCell.CELL_TYPE_NUMERIC){
					SimpleDateFormat sdf;
					if (cell.getCellStyle().getDataFormat() == HSSFDataFormat  
	                        .getBuiltinFormat("h:mm")) {  
	                    sdf = new SimpleDateFormat("HH:mm");  
	                } else {// 日期  
	                    sdf = new SimpleDateFormat("yyyy-MM-dd");  
	                }  
	                java.util.Date date =  cell.getDateCellValue();  
	                value = sdf.format(date);  
				}else{
				}*/
				value = cell.toString().trim();
				//去掉一些乱七八糟的空格
				//Speicher [%1$s /% 2 $ s] Speicher [% 3 $ s /% 4 $ s]
				
				
				
				value = value.replaceAll("% ", "%")
							.replaceAll("1 ", "1").replaceAll("2 ", "2")
							.replaceAll("3 ", "3").replaceAll("4 ", "4")
							.replaceAll(" s", "s").replaceAll(" d", "d")
							.replaceAll("％", "%").replace("\'", "\\'");
							
					
				
				
				//System.out.println("system: "+value);
				
				File file = new File(dir, mBuilder.getStringName());
				if (file.exists()) {
					FileOutputStream fos = null;
					try {
						fos = new FileOutputStream(file,true);
						StringBuilder builder = new StringBuilder();
						builder.append("\t")
								// 有个小空格
								.append("<string name=\"").append(mStringKey)
								.append("\">").append(value).append("</string>")
								.append("\r\n");
	
						if ((MAX_ROW - mBuilder.getIgnoreRow()) == rowIndex) { 
																	
							builder.append("</resources>\r\n");
						}
						
						fos.write(builder.toString().getBytes("utf-8"));
					} catch (Exception e) {
						e.printStackTrace();
					} finally {
						if (fos != null) {
							try {
								fos.close();
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
						}
					}
				}
			}
		}

	}
	
	
	
	/**
	 * 写数据到 array 中
	 * @param cIndex 列
	 * @param cell 单元格
	 * @param rowIndex 行
	 */
	public void writeValueToArray(int cIndex,Cell cell,int rowIndex){
		String value = "";
		StringBuilder sb = null;
		if(cell != null){
			value = cell.toString();
			//System.out.println("value:  "+value);
			value = value.replaceAll(" ", "");
			//System.out.println("value:  "+value);
			if(cIndex == 0){  //第一列
				//第一次我们先把这个
				sb = new StringBuilder();
				
				if(!isLastItemString){
					sb.append("\t")
					  .append("<string-array name =\"").append(value).append("\">");
					writeDatatoArrayFile(sb.toString(),null);
					
				}else{
					
					//补全上一次
					sb = new StringBuilder();
					sb.append("\n")
					  .append("\t")
					  .append("</string-array>");
					writeDatatoArrayFile(sb.toString(),null);
					//开始这一次
					sb = new StringBuilder();
					sb.append("\n\n")
					  .append("\t")
					  .append("<string-array name =\"").append(value).append("\">");
					writeDatatoArrayFile(sb.toString(),null);
					isLastItemString = false;
				}
			}else if(cIndex>1){
				//写数据
				sb = new StringBuilder();
				sb.append("\n")
				  .append("\t\t")
				  .append("<item>")
				  .append(value)
				  .append("</item>");
				
				
				isLastItemString = true;
				 	
				writeDatatoArrayFile(sb.toString(),mMap.get(cIndex));
				
				if(rowIndex == MAX_ROW - mBuilder.getIgnoreRow()
						&& cIndex == MAX_COLUME - 1){
					sb = new StringBuilder();
					sb.append("\n")
					  .append("\t")
					  .append("</string-array>\r\n")
					  .append("</resources>\r\n");
					writeDatatoArrayFile(sb.toString(),null);
				}
			}
		}
	}
	
	
	/**
	 * 写数据到 array file 里面
	 * @param value
	 * @param isLoop ，loop 主要用于array开头的key
	 */
	private  void writeDatatoArrayFile(String value,String dir){
		//array的字符串比较麻烦，自己定义的格式，从第二列开始
		List<File> files = new ArrayList<>();
		if(dir != null){
			File file = new File(dir,mBuilder.getArrayName());
			files.add(file);
			
		}else{
			for (int i = 2; i < MAX_COLUME; i++) {
				//把能获取到文件夹的列，写入 array 字符串
				String path = mMap.get(i);
				File file = new File(path,mBuilder.getArrayName());
				files.add(file);
				
			}
		}
		
		for (File file : files) {
			if (file.exists()) {
				FileOutputStream fos = null;
				
				try {
					fos = new FileOutputStream(file,true);
					fos.write(value.getBytes("utf-8"));
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					if (fos != null) {
						try {
							fos.close();
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}
				}
			}
		}
		
	}

}
