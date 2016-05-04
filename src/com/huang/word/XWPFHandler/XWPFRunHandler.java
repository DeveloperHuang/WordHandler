package com.huang.word.XWPFHandler;

import javax.xml.namespace.QName;

import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class XWPFRunHandler {

	/**
	 * 获取所有的文本内容
	 * @param run
	 * @return
	 */
	public static String getText(XWPFRun run) {
		CTR ctr = run.getCTR();
		int tArraySize = ctr.sizeOfTArray();
		if (tArraySize == 0) {
			return null;
		} else {
			StringBuffer text = new StringBuffer();
			for (int i = 0; i < tArraySize; i++) {
				text.append(ctr.getTArray(i).getStringValue());
			}
			return text.toString();
		}
	}
	
	/**
	 * 重新设置XWPFRun的文本内容
	 * @param run
	 * @param text
	 */
	public static void setText(XWPFRun run,String text){
		
		CTR ctr = run.getCTR();
		CTText textArray[] = getTextArray(ctr);
		if(textArray.length <= 0){
			setCTText(ctr.addNewT(),text);
		}else{
			setCTText(textArray[0],text);
			for (int i = 1; i < textArray.length; i++) {
				setCTText(textArray[i],"");
			}
		}
	}
	
	private static CTText[] getTextArray(CTR ctr) {
		int tArraySize = ctr.sizeOfTArray();
		CTText[] textArray = new CTText[tArraySize];
		for (int i = 0; i < tArraySize; i++) {
			textArray[i] = ctr.getTArray(i);
		}
		return textArray;
	}
	
	private static void setCTText(CTText cttext,String text){
		cttext.setStringValue(text);
		preserveSpaces(cttext);
	}
	
	private static void preserveSpaces(XmlString xs) {
        String text = xs.getStringValue();
        if (text != null && (text.startsWith(" ") || text.endsWith(" "))) {
            XmlCursor c = xs.newCursor();
            c.toNextToken();
            c.insertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace", "space"), "preserve");
            c.dispose();
        }
    }
}
