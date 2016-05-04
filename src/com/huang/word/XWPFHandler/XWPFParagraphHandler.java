package com.huang.word.XWPFHandler;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class XWPFParagraphHandler {
	
	private XWPFParagraph paragraph;
	
	private List<XWPFRun> allXWPFRuns;
	//所有run的String合并后的内容
	private StringBuffer context ;
	//长度与context对应的RunChar集合
	List<RunChar> runChars ;
	
	public XWPFParagraphHandler(XWPFParagraph paragraph){
		this.paragraph = paragraph;
		initParameter();
	}
	
	/**
	 * 初始化各参数
	 */
	private void initParameter(){
		context = new StringBuffer();
		runChars = new ArrayList<XWPFParagraphHandler.RunChar>();
		allXWPFRuns = new ArrayList<XWPFRun>();
		setXWPFRun();
	}
	
	
	/**
	 * 设置XWPFRun相关的参数
	 * @param run
	 * @throws Exception
	 */
	private void setXWPFRun() {
		
		allXWPFRuns = paragraph.getRuns();
		if(allXWPFRuns == null || allXWPFRuns.size() == 0){
			return;
		}else{
			for (XWPFRun run : allXWPFRuns) {
				int testPosition = run.getTextPosition();
				String text = run.getText(testPosition);
				if(text == null || text.length() == 0){
					return;
				}
				
				this.context.append(text);
				for(int i = 0 ; i < text.length() ; i++){
					runChars.add(new RunChar(text.charAt(i), run));
				}
			}
		}
		System.out.println(context.toString());
	}
	
	/**
	 * 获取所有的文本内容
	 * @return
	 */
	public String getString(){
		return context.toString();
	}
	
	/**
	 * 判断是否包含指定的内容
	 * @param key
	 * @return
	 */
	public boolean contains(String key){
		return context.indexOf(key) >= 0 ? true : false;
	}
	
	/**
	 * 所有匹配的值替换为对应的值
	 * @param key(匹配模板中的${key})
	 * @param value 替换后的值
	 * @return
	 */
	public boolean replaceAll(String key,String value){
		boolean replaceSuccess = false;
		key = "${" + key + "}";
		while(replace(key, value)){
			replaceSuccess = true;
		}
		return replaceSuccess;
	}
	
	/**
	 * 所有匹配的值替换为对应的值(key匹配模板中的${key})
	 * @param param 要替换的key-value集合
	 * @return
	 */
	public boolean replaceAll(Map<String,String> param){
		Set<Entry<String, String>> entrys = param.entrySet();
		boolean replaceSuccess = false;
		for (Entry<String, String> entry : entrys) {
			String key = entry.getKey();
			boolean currSuccessReplace = replaceAll(key,entry.getValue());
			replaceSuccess = replaceSuccess?replaceSuccess:currSuccessReplace;
		}
		return replaceSuccess;
	}
	
	/**
	 * 将第一个匹配到的值替換为对应的值
	 * @param key 
	 * @param value
	 * @return
	 */
	private boolean replace(String key,String value){
		if(contains(key)){
			/*
			 * 1:得带key对应的开始和结束下标
			 */
			int startIndex = context.indexOf(key);
			int endIndex = startIndex+key.length();
			/*
			 * 2:获取第一个匹配的XWPFRun
			 */
			RunChar startRunChar = runChars.get(startIndex);
			XWPFRun startRun = startRunChar.getRun();
			/*
			 * 3:将匹配的key清空
			 */
			runChars.subList(startIndex, endIndex).clear();
			/*
			 * 4:将value设置到startRun中
			 */
			List<RunChar> addRunChar = new ArrayList<XWPFParagraphHandler.RunChar>();
			for(int i = 0 ; i < value.length() ; i++){
				addRunChar.add(new RunChar(value.charAt(i), startRun));
			}
			runChars.addAll(startIndex, addRunChar);
			resetRunContext(runChars);
			return true;
		}else{
			return false;
		}
	}
	
	/**
	 * 重新设置公共的参数
	 * @param newRunChars
	 */
	private void resetRunContext(List<RunChar> newRunChars){
		/*
		 * 生成新的XWPFRun与Context的对应关系
		 */
		HashMap<XWPFRun, StringBuffer> newRunContext = new HashMap<XWPFRun, StringBuffer>();
		//重设context
		context = new StringBuffer();
		for(RunChar runChar : newRunChars){
			StringBuffer newRunText ;
			if(newRunContext.containsKey(runChar.getRun())){
				newRunText = newRunContext.get(runChar.getRun());
			}else{
				newRunText = new StringBuffer();
			}
			context.append(runChar.getValue());
			newRunText.append(runChar.getValue());
			newRunContext.put(runChar.getRun(), newRunText);
		}
		
		/*
		 * 遍历旧的runContext,替换context
		 * 并重新设置run的text,如果不匹配,text设置为""
		 */
		for(XWPFRun run : allXWPFRuns){
			if(newRunContext.containsKey(run)){
				String newContext = newRunContext.get(run).toString();
				XWPFRunHandler.setText(run,newContext);
			}else{
				XWPFRunHandler.setText(run,"");
			}
		}
	}
	
	/**
	 * 实体类:存储字节与XWPFRun对象的对应关系
	 * @author JianQiu
	 */
	class RunChar{
		/**
		 * 字节
		 */
		private char value;
		/**
		 * 对应的XWPFRun
		 */
		private XWPFRun run;
		public RunChar(char value,XWPFRun run){
			this.setValue(value);
			this.setRun(run);
		}
		public char getValue() {
			return value;
		}
		public void setValue(char value) {
			this.value = value;
		}
		public XWPFRun getRun() {
			return run;
		}
		public void setRun(XWPFRun run) {
			this.run = run;
		}
		
	}
}
