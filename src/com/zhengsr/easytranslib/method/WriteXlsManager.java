package com.zhengsr.easytranslib.method;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import com.zhengsr.easytranslib.bean.CusRow;
import com.zhengsr.easytranslib.bean.FloderBean;
import com.zhengsr.easytranslib.bean.XlsWriteBean;


/**
 * @author zhengshaorui 2018/6/24
 */
public class WriteXlsManager extends BaseSheet{
	
	private static final String ARRAY_TYPE_DECARE = "type_array(不要修改这个申明)";
	private static final String STRING_TYPE_DECARE = "type_string(不要修改这个申明)";
	private static final String STRING_NAME = "strings";
	private static final String ARRAY_NAME = "arrays";
	private boolean isArrayXls = false;
	private List<String> mKeyList = new ArrayList<>(); //用来string的key保存
	private Workbook mWorkbook;
	private static class Holder{
		static WriteXlsManager writeXlsManager = new WriteXlsManager();
	}
	public static WriteXlsManager getInstance(){
		return Holder.writeXlsManager;
	}
	
	private WriteXlsManager(){
		
	}
	
	
	
	public void startWrite(XlsWriteBean.Builder builder){
		File file = new File(builder.getRootPath(),builder.getFileFloderName());
		FileOutputStream fos =  null;
		if(file.exists()){
			try {
				//创建一个工作簿
				if(builder.getXlsName().toLowerCase().endsWith("xlsx")){
					mWorkbook = new XSSFWorkbook();
				}else{
					mWorkbook = new HSSFWorkbook();
				}
			  //如果已经有了，先删除
				String filePath = builder.getRootPath()+File.separator+builder.getFileFloderName();
			    File xlsFile = new File(filePath,builder.getXlsName());
			    if(xlsFile.exists()){
			    	xlsFile.delete();
			    }
			    //开始写数据到 xls 
			    startWriteWorkbook(file);
			    
			  
	        	fos = new FileOutputStream(new File(filePath,builder.getXlsName()));
				mWorkbook.write(fos);
				fos.close();
			} catch (Exception e) {
				// TODO: handle exception
				System.out.println("error: "+e.toString());
			}finally{
				if(fos != null){
					try {
						fos.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		}else{
			System.out.println("cannot find "+builder.getRootPath()+File.separator+" "+builder.getFileFloderName());
		}
	}
	
	/**
	 * 开始写数据，主要配置第一行，还有写相关数据
	 * @param rootFile
	 */
	private void startWriteWorkbook(File rootFile){
		if (rootFile.exists()) {
			//设置自动换行
		    CellStyle cellStyle= mWorkbook.createCellStyle();     
		    cellStyle.setWrapText(true);   
		    CreationHelper createHelper = mWorkbook.getCreationHelper();
			FloderBean bean = getFloderBean(rootFile); //获取到所有路径,并保存到 floderBean 中
			if (bean != null) {
				mWorkbook.createSheet();
				if(!bean.floderPaths.isEmpty()){
					
					for (String floderPath : bean.floderPaths) { //浏览文件夹
						File floderFile = new File(floderPath);
						if(floderFile.exists()){
							
							
							List<String> valueNames = getFloderFileNameList(floderPath);
							for (String valueName : valueNames) { //浏览各个文件夹中的文件
								System.out.println("valuename: "+valueName);
								Sheet sheet = null;
								Row row = null;
								int startColume = 1;
								int size = valueNames.size();
								if(valueName.contains("array")){
									isArrayXls = true;
									if(size == 1){ //如果只有 array，那么sheet 从0开始
										mWorkbook.setSheetName(0, ARRAY_NAME);
										sheet = mWorkbook.getSheet(ARRAY_NAME);
									}else{ //如果有string，则自动排到第二个
										sheet = mWorkbook.getSheet(ARRAY_NAME);
										if(sheet == null){
											sheet = mWorkbook.createSheet(ARRAY_NAME);
										}
									}
									
									row = sheet.createRow(0);
									row.createCell(0).setCellValue(ARRAY_TYPE_DECARE);
									
									
								}else{
									isArrayXls = false;
									
									sheet = mWorkbook.getSheet(STRING_NAME);
									if(sheet == null){
										mWorkbook.setSheetName(0, STRING_NAME);
										sheet = mWorkbook.getSheet(STRING_NAME);
									}
									
									row = sheet.createRow(0);
									row.createCell(0).setCellValue(STRING_TYPE_DECARE);
									
								}
								sheet.setColumnWidth(0, 30*256);
								if(isArrayXls){
									startColume = 2;
								}
								//开始写其他行的数据
								for (int langIndex = 0; langIndex < bean.languages.size(); langIndex++) {
									row.createCell(langIndex+startColume).setCellValue(
											createHelper.createRichTextString(bean.languages.get(langIndex)));
									List<CusRow> lists = parseStringXml(bean.floderPaths.get(langIndex),
											valueName,createHelper);
									if(lists != null && !lists.isEmpty()){
										writeDataToXls(lists,sheet,langIndex);
									}
								}
							}
						}
					}
				}
			}
			
		
		}
	}
	
	
	
	
	/**
	 * 获取路径下数据
	 * @param rootFile
	 * @return
	 */
	private FloderBean getFloderBean(File rootFile){
		FloderBean bean = new FloderBean();
		bean.path = rootFile.getAbsolutePath();
		bean.languages = new ArrayList<>();
		bean.floderPaths = new ArrayList<>();
		if(rootFile.exists()){
			//bean.floderNames.add(getFloderName(file));
			File[] files = rootFile.listFiles();
			for (File file : files) {
				bean.languages.add(getLangByFloder(getFloderName(file)));
				bean.floderPaths.add(file.getAbsolutePath());
			}
			
		}
		
		return bean;
		
	}
	
	
	
	/**
	 * 写数据到 xls 中
	 * @param lists
	 * @param sheet
	 * @param rIndex
	 */
	private void writeDataToXls(List<CusRow> lists,Sheet sheet,int rIndex){
		
			List<Integer> itemCountList = new ArrayList<>(); //记录item的个数
			CellStyle cellStyle= mWorkbook.createCellStyle();     
		    cellStyle.setWrapText(true); 
			int size = lists.size();
				//列
				for (int cIndex = 0; cIndex < size; cIndex++) {
					CusRow cusRow = lists.get(cIndex);
					/**
					 * 把 string 的数据，写进 xls 中
					 * 这里的原理是：
					 * 1、当cIndex == 0,即当前为value文件夹，以这个为基准，保存 key 和 row ，row 用map保存，记录第几行
					 * 2、当cIndex != 0,即其他文件夹，但是又不能保证 strings.xml 和 values 里的 strings.xml 是一致的
					 * 	 所以，通过第一次保存的 key，获取到对应的 index，然后通过这个index，即可获取到对应的 row，这样就不会重复了
					 */
					if(!isArrayXls){ //string
						if(rIndex == 0){
							//从第二行开始
							Row valueRow = sheet.createRow(cIndex+1);
							valueRow.createCell(0).setCellValue(cusRow.key);//先写key
							Cell cell = valueRow.createCell(rIndex+1);
							cell.setCellStyle(cellStyle);
							sheet.setColumnWidth(cIndex+1, 25*256);
							cell.setCellValue(cusRow.value); //再写value
							mKeyList.add(cusRow.key);
						}else{ //array
							int index = getIndexFromKey(cusRow.key);
							if(index != -1){
								//找到对应的row
								Row valueRow = sheet.getRow(index+1);
								if(sheet == null){
									valueRow = sheet.createRow(index+1);
								}
								
								Cell cell = valueRow.createCell(rIndex+1);
								cell.setCellStyle(cellStyle);
								sheet.setColumnWidth(index, 35*256);
								cell.setCellValue(cusRow.value);
							}
							
						}
					}else{
						/**
						 * 把 array 的数据，写进 xls 中
						 * 这里的原理是(先看生成的表格再来看这段思路)：
						 * 1、首先先写 key 和 item，key在第一行，第一列；所以，当有数据过来的时候，我们首先，先判断是否是 key，如果是，
						 * 	  则写上 key，并继续判断是否有 item，如果有，则把item逐行写上；
						 * 2、当 item 写完，又遇到key，这时候，行的起始位置，应该是上一次的  key 加 item 的行数 再加1，补充标签
						 */
						int index = cIndex + 1; 
						if(cIndex > 0){
							int num = 0;
							for (int itemIndex = 0; itemIndex < itemCountList.size(); itemIndex++) {
								num += itemCountList.get(itemIndex);
							}
							index = cIndex + 1 + num;
						}
							Row valueRow = sheet.createRow(index);
							if(cusRow.key != null){
								valueRow.createCell(0).setCellValue(cusRow.key);
								if(cusRow.items != null){
									int count = cusRow.items.size() - 1;
									itemCountList.add(count);
									index = index + 1;
									System.out.println("count: "+count+" ");
									for (int cRow = 0; cRow < count; cRow++) {
										String item = cusRow.items.get(cRow);
										int itemIndex = index + cRow ;
										Row itemRow = null;
										itemRow = sheet.getRow(itemIndex);
										if(itemRow == null){
											itemRow = sheet.createRow(itemIndex);
										
										}
										itemRow.createCell(1).setCellValue("<item>");
										Cell cell = itemRow.createCell(rIndex + 2); //从第二列开始
										cell.setCellStyle(cellStyle);
										sheet.setColumnWidth(cIndex+1, 35*256);
										cell.setCellValue(item);
										System.out.println("item: "+itemRow.getCell(0)+" "+itemRow.getCell(1));
									}
								}
							}
					}
				}
			
	
	}
	
	
	/**
	 * 使用 DOM 解析 xml
	 * @param path
	 * @param stringName
	 * @param createHelper
	 * @return
	 */
	private List<CusRow> parseStringXml(String path,String stringName,CreationHelper createHelper) {
		List<CusRow> lists = new ArrayList<>();
        try {
        	File file = new File(path,stringName); 
        	if (file.exists()) {
        		SAXReader read = new SAXReader();
        		org.dom4j.Document document =  read.read(file);
        		 Element root = document.getRootElement();//获取根元素
        	  //   List<Element> childElements = root.elements();//获取当前元素下的全部子元素
        	
        		 
    		    for (Iterator<Element> it = root.elementIterator(); it.hasNext();) {
    		        Element element = it.next();
    		        // do something
    		        if(element != null){
	    		      //  System.out.println("test: "+element.attributeValue("name")+" "+element.getStringValue());
	    		        CusRow cusRow = new CusRow();
	    		        cusRow.key = element.attributeValue("name");
	    		        cusRow.value = element.getStringValue();
	    		        if(isArrayXls){
		                	 if(cusRow.value != null){
		                		 String[] items = cusRow.value.split("\n");
		                		 if(items != null){
		                			 cusRow.items = new ArrayList<>();
		                			 for (String item : items) {
										if(item != null && item.length() > 1){
											cusRow.items.add(item);
										}
									}
		                		 }
		                	 }
	               	 	}
	    		        lists.add(cusRow);
    		        }
    		    }
			}
           
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("parseStringXml error: "+e.toString());
        }
        return lists;
    }
	
	/**
	 * 获取文件夹的名字
	 * @param file
	 * @return
	 */
	private String getFloderName(File file){
		String path = file.getAbsolutePath();
		//window 路径转换
		String[] paths = path.replace("\\", "/").toString().split("/");
		if(paths != null && paths.length > 0){
			return paths[paths.length - 1];
		}
		return null;
	}
	
	/**
	 * 获取文件夹中的文件名
	 * @param path
	 * @return
	 */
	private List<String> getFloderFileNameList(String path){
		File dir = new File(path);
		List<String> lists = new ArrayList<>();
		if (dir.exists()) {
			File[] files = dir.listFiles();
			if(files != null){
				int length = files.length;
				for (int i = 0; i < length; i++) {
					File file = files[i];
					String[] paths = file.getAbsolutePath().toString()
							.replace("\\", "/").toString().split("/");
					//return paths[paths.length - 1];
					lists.add(paths[paths.length - 1]);
				
				}
			}
		}
		return lists;
	}
	

	/**
	 * 获取key中的 index
	 * @param key
	 * @return
	 */
	private int getIndexFromKey(String key){
		if(!mKeyList.isEmpty()){
			for (String rowkey : mKeyList) {
				if (key.equals(rowkey)) {
					return mKeyList.indexOf(key);
				}
			}
		}
		return -1;
	}

}
