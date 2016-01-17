package com.cm.oe.word_table_test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import com.cm.oe.test.UpdateBudget3;

public class XwpfTestBudget3 {
	private final static String filePath = "testfiles/templete_all.docx";
	String path1 = "testfiles/ysb_final.xls";
	String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private UpdateBudget3 ub = new UpdateBudget3(path1, path2);
	/**
	 * 通过XWPFDocument对内容进行访问。对于XWPF文档而言，用这种方式进行读操作更佳。
	 * 
	 * @throws Exception
	 */
	@Test
	public void testReadByDoc() throws Exception {
		InputStream is = new FileInputStream(filePath);
		OutputStream os = new FileOutputStream("testfiles/templete_all2.docx");
		XWPFDocument doc = new XWPFDocument(is);
		String zh = ub.getZhFrom4GYsb(path1, path2);
		List<String> datas = ub.readExcel(zh);
		// 获取文档中所有的表格
		List<XWPFTable> tables = doc.getTables();

		XWPFTable table = tables.get(8);
		CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
		genBorders(borders);
		XWPFTableRow tableOneRowTwo = table.createRow();
		tableOneRowTwo.setHeight(11);
		int i = 0;
		for(String values:datas){
			setCellText(tableOneRowTwo.getCell(i), values, "FFFFFF", 21);
			i++;
		}
		doc.write(os);
		os.flush();
		os.close();
		this.close(is);
		doc.close();
	}
 
	  // 设置单元格文字  
    public void setRowCellText(XWPFTableCell cell, String text, int width,  
            boolean isShd, int shdValue, String shdColor, STVerticalJc.Enum jc,  
            STJc.Enum stJc) {  
        CTTc cttc = cell.getCTTc();  
        CTTcPr ctPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();  
        CTShd ctshd = ctPr.isSetShd() ? ctPr.getShd() : ctPr.addNewShd();  
        if (isShd) {  
            if (shdValue > 0 && shdValue <= 38) {  
                ctshd.setVal(STShd.Enum.forInt(shdValue));  
            }  
            if (shdColor != null) {  
                ctshd.setColor(shdColor);  
            }  
        }  
        ctPr.addNewVAlign().setVal(jc);  
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(stJc); 
        XWPFParagraph cellP=cell.getParagraphs().get(0);
		XWPFRun cellR = cellP.createRun();
		cellR.setFontSize(11);
		cellR.setText(text);
    }  
    
	private void genBorders(CTTblBorders borders) {
		CTBorder hBorder = borders.addNewInsideH();
		hBorder.setVal(STBorder.Enum.forString("thick"));
		hBorder.setSz(new BigInteger("1"));
		hBorder.setColor("000000");
		//
		CTBorder vBorder = borders.addNewInsideV();
		vBorder.setVal(STBorder.Enum.forString("thick"));
		vBorder.setSz(new BigInteger("1"));
		vBorder.setColor("000000");
		//
		CTBorder lBorder = borders.addNewLeft();
		lBorder.setVal(STBorder.Enum.forString("thick"));
		lBorder.setSz(new BigInteger("1"));
		lBorder.setColor("000000");
		//
		CTBorder rBorder = borders.addNewRight();
		rBorder.setVal(STBorder.Enum.forString("thick"));
		rBorder.setSz(new BigInteger("1"));
		rBorder.setColor("000000");
		//
		CTBorder tBorder = borders.addNewTop();
		tBorder.setVal(STBorder.Enum.forString("thick"));
		tBorder.setSz(new BigInteger("1"));
		tBorder.setColor("000000");
		//
		CTBorder bBorder = borders.addNewBottom();
		bBorder.setVal(STBorder.Enum.forString("thick"));
		bBorder.setSz(new BigInteger("1"));
		bBorder.setColor("000000");
	}

	public void setCellText(XWPFTableCell cell, String text, String bgcolor,
			int width) {
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
        XWPFParagraph cellP=cell.getParagraphs().get(0);
        cellP.setAlignment(ParagraphAlignment.LEFT);
        cellP.setVerticalAlignment(TextAlignment.AUTO);
		XWPFRun cellR = cellP.createRun();
		cellR.setFontSize(10);
		cellR.setText(text.trim());
	}

	public void addNewPage(XWPFDocument document,BreakType breakType){
		XWPFParagraph xp = document.createParagraph();
		xp.createRun().addBreak(breakType);
	}
	
	/**
	 * 关闭输入流
	 * 
	 * @param is
	 */
	private void close(InputStream is) {
		if (is != null) {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {
		try {
			XwpfTestBudget3 xwpf = new XwpfTestBudget3();
			xwpf.testReadByDoc();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}