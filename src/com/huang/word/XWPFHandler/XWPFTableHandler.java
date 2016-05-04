package com.huang.word.XWPFHandler;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * 用于处理XWPFTable字符替换的工具类
 * @author JianQiu
 */
public class XWPFTableHandler {
	
	private XWPFTable xwpfTable;
	private List<XWPFTableRow> xwpfTableRows;
	
	/**
	 * 初始化XWPFTable处理器
	 * @param xwpfTable 待处理的对象
	 */
	public XWPFTableHandler(XWPFTable xwpfTable){
		this.xwpfTable = xwpfTable;
		xwpfTableRows = xwpfTable.getRows();
	}
	
	/**
	 * 获取所有的行
	 * @return 包含了行对象的集合
	 */
	public List<XWPFTableRow> getRows(){
		return xwpfTableRows;
	}
	
	/**
	 * 获取所有的cell
	 * @return 包含了cell的集合
	 */
	public List<XWPFTableCell> getCells(){
		List<XWPFTableCell> xwpfTableCells = new ArrayList<XWPFTableCell>();
		for(int i = 0 ; i < xwpfTableRows.size() ; i++){
			xwpfTableCells.addAll(getCells(i));
		}
		return xwpfTableCells;
	}
	
	/**
	 * 得到table中指定行对应的cell集合
	 * @param rowIndex 行数(从0开始)
	 * @return 包含了指定行cell的集合
	 */
	public List<XWPFTableCell> getCells(int rowIndex){
		return xwpfTableRows.get(rowIndex).getTableCells();
	}
	
	/**
	 * 得到table中所有的文本内容
	 * @return 
	 */
	public List<XWPFParagraph> getAllParagraphs(){
		List<XWPFTableCell> xwpfTableCells = getCells();
		List<XWPFParagraph> xwpfParagraphs = new ArrayList<XWPFParagraph>();
		for(XWPFTableCell cell : xwpfTableCells){
			xwpfParagraphs.addAll(cell.getParagraphs());
		}
		return xwpfParagraphs;
	}
	
	/**
	 * 所有匹配的值替换为对应的值
	 * @param key(匹配模板中的${key})
	 * @param value 替换后的值
	 * @return
	 */
	public boolean replace(String key,String value){
		List<XWPFParagraph> allParagraphs = getAllParagraphs();
		boolean successReplace = false;
		for(XWPFParagraph xwpfParagraph : allParagraphs){
			XWPFParagraphHandler xwpfParagraphUtils = new XWPFParagraphHandler(xwpfParagraph);
			boolean currSuccessTag = xwpfParagraphUtils.replaceAll(key, value);
			successReplace = successReplace?successReplace:currSuccessTag;
		}
		return successReplace;
	}
	
	/**
	 * 所有匹配的值替换为对应的值(key匹配模板中的${key})
	 * @param param 要替换的key-value集合
	 * @return
	 */
	public boolean replace(Map<String,String> param){
		List<XWPFParagraph> allParagraphs = getAllParagraphs();
		boolean successReplace = false;
		for(XWPFParagraph xwpfParagraph : allParagraphs){
			XWPFParagraphHandler xwpfParagraphUtils = new XWPFParagraphHandler(xwpfParagraph);
			boolean currSuccessTag = xwpfParagraphUtils.replaceAll(param);
			successReplace = successReplace?successReplace:currSuccessTag;
		}
		return successReplace;
	}
	
}
