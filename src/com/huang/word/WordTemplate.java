package com.huang.word;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.huang.word.XWPFHandler.XWPFParagraphHandler;
import com.huang.word.XWPFHandler.XWPFTableHandler;

/**
 * 仅支持对docx文件的文本及表格中的内容进行替换
 * 模板仅支持 ${key} 标签
 * @author JianQiu
 *
 */
public class WordTemplate {
	
	private XWPFDocument document;
	
	/**
	 * 初始化模板内容
	 * @param inputStream 模板的读取流(docx文件)
	 * @throws IOException
	 */
	public WordTemplate(InputStream inputStream) throws IOException{
		document = new XWPFDocument(inputStream);
	}
	
	/**
	 * 替换模板中的标签为实际的内容
	 * @param map 
	 */
	public void replaceTag(Map<String,String> map){
		replaceParagraphs(map);
		replaceTables(map);
	}
	
	/**
	 * 将处理后的内容写入到输出流中
	 * @param outputStream 输出流
	 * @throws IOException
	 */
	public void write(OutputStream outputStream) throws IOException{
		document.write(outputStream);
	}
	
	/**
	 * 替换文本中的标签
	 * @param map key(待替换标签)-value(文本内容)
	 */
	private void replaceParagraphs(Map<String,String> map){
		List<XWPFParagraph> allXWPFParagraphs = document.getParagraphs();
		for (XWPFParagraph XwpfParagrapg : allXWPFParagraphs) {
			XWPFParagraphHandler XwpfParagrapgUtils = new XWPFParagraphHandler(XwpfParagrapg);
			XwpfParagrapgUtils.replaceAll(map);
		}
	}
	
	/**
	 * 替换表格中的标签 
	 * @param map key(待替换标签)-value(文本内容)
	 */
	private void replaceTables(Map<String,String> map){
		List<XWPFTable> xwpfTables = document.getTables();
		for(XWPFTable xwpfTable : xwpfTables){
			XWPFTableHandler xwpfTableUtils = new XWPFTableHandler(xwpfTable);
			xwpfTableUtils.replace(map);
		}
	}
	
}
