package com.zhengsr.easytranslib.bean;
/**
 * @author zhengshaorui 2018/6/24
 */
public class XlsReadBean {
	private Builder mBuilder;
	public XlsReadBean(Builder builder){
		mBuilder = builder;
	}
	
	public XlsReadBean.Builder getBuilder(){
		return mBuilder;
	}
	
	
	public static class Builder{
		String rootPath;
		String xlsFile;
		String fileFloderName; 
		int ignoreRow;
		String stringName = "strings.xml";
		String arrayName = "array.xml"; 
		
		public Builder setRootPath(String rootPath) {
			this.rootPath = rootPath;
			return this;
		}
		
		public Builder setFileFloderName(String fileFloderName) {
			this.fileFloderName = fileFloderName;
			return this;
		}
		
		public Builder setXlsFile(String xlsFile){
			this.xlsFile = xlsFile;
			return this;
		}
		
		public Builder setIgnoreRow(int ignoreRow) {
			this.ignoreRow = ignoreRow;
			return this;
		}
		
		public Builder setStringName(String stringName) {
			this.stringName = stringName;
			return this;
		}
		
		public Builder setArrayName(String arrayName) {
			this.arrayName = arrayName;
			return this;
		}
		public XlsReadBean builder(){
			checkNull(this);
			return new XlsReadBean(this);
		}

		//get 方法
		public String getRootPath() {
			return rootPath;
		}

		public String getFileFloderName() {
			return fileFloderName;
		}

		public int getIgnoreRow() {
			return ignoreRow;
		}

		public String getStringName() {
			return stringName;
		}

		public String getArrayName() {
			return arrayName;
		}
		
		public String getXlsFile(){
			return xlsFile;
		}
		
		
		
		
	}


	private static void checkNull(Builder builder) {
		// TODO Auto-generated method stub
		if(builder.getRootPath() == null){
			throw new NullPointerException("you need to set root path!");
		}
		if(builder.getXlsFile() == null){
			throw new NullPointerException("you need to set xlsfile file!");
		}
		if(builder.getFileFloderName() == null){
			throw new NullPointerException("you need to set file name!");
		}
		
	}
}
