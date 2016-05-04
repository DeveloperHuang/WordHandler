package com.huang.word;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.BeforeClass;
import org.junit.Test;

import com.huang.word.XWPFHandler.XWPFParagraphHandler;
import com.huang.word.XWPFHandler.XWPFTableHandler;

public class TestWordTemplate {
	
	private static WordTemplate template;
	private static final String DOCX_MODEL_PATH = "doc/model.docx";
	private static final String DOCX_FILE_WRITE = "doc/replaceModelWord.docx";
	private static Map<String,String> map = new HashMap<String, String>();
	
	@BeforeClass
	public static void init(){
		File file = new File(DOCX_MODEL_PATH);
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(file);
			template = new WordTemplate(fileInputStream);
		} catch(IOException exception){
			exception.printStackTrace();
		}
		
		//设置参数
		map.put("title", "docx模板生成的内容");
		map.put("user1", "张三");
		map.put("user2", "李四");
		map.put("text1", "上海");
		map.put("text2", "广州");
		map.put("cell1", "1行1列");
		map.put("cell2", "1行2列");
		map.put("cell3", "2行1列");
		map.put("cell4", "2行2列");
		map.put("year", "2016");
		map.put("month", "04");
		map.put("day", "03");
	}
	
	public void testReplaceTag(){
		template.replaceTag(map);
	}
	
	@Test
	public void testWrite(){
		testReplaceTag();
		File file = new File(DOCX_FILE_WRITE);
		FileOutputStream out;
		try {
			out = new FileOutputStream(file);
			BufferedOutputStream bos = new BufferedOutputStream(out);
			template.write(bos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
}
