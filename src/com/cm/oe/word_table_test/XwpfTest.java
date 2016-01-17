package com.cm.oe.word_table_test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import com.cm.oe.test.UpdateBudget;

public class XwpfTest {
	private final static String filePath = "testfiles/templete_all.docx";
	private UpdateBudget ub = new UpdateBudget();
	String path1 = "testfiles/ysb_final.xls";
	String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
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
		Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
		Map<String, List<String>> data_b3 = ub.getB3(path1);
		Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path1, path2);
		// List<XWPFParagraph> paras = doc.getParagraphs();
		// for (XWPFParagraph para : paras) {
		// //当前段落的属性
		// // CTPPr pr = para.getCTP().getPPr();
		// System.out.println(para.getText());
		// }
		// 获取文档中所有的表格
		List<XWPFTable> tables = doc.getTables();
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;

		XWPFTable table = tables.get(2);
		CTTbl ttbl = table.getCTTbl();
		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
//		tblPr.addNewTblStyle()
		// CTTblBorders borders=tblPr.getTblBorders();
		CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
		genBorders(borders);
//		rows = table.getRows();
//		for (XWPFTableRow row : rows) {
//			// 获取行对应的单元格
//			CTRow ctrpr = row.getCtRow();
//			cells = row.getTableCells();
//			for (XWPFTableCell cell : cells) {
//				CTTc cttc = cell.getCTTc();
////				CTTcPr cttcpr = cttc.getTcPr();
////				CTCnf cttt = cttc.getTcPr().getCnfStyle();
//				System.out.println(cell.getText());
//			}
//		}
		for(String key:datas_map.keySet()){
			XWPFTableRow tableOneRowTwo = table.createRow();
			tableOneRowTwo.setHeight(11);
			setCellText(tableOneRowTwo.getCell(0), key, "FFFFFF", 21);
			setCellText(tableOneRowTwo.getCell(1), datas_map.get(key).get(0), "FFFFFF", 21);
			setCellText(tableOneRowTwo.getCell(2), datas_map.get(key).get(1), "FFFFFF", 21);
		}

//		tableOneRowTwo.getCell(0).setText("测试");
//		tableOneRowTwo.getCell(1).setText("测试");
//		tableOneRowTwo.getCell(2).setText("测试");


//		setRowCellText(tableOneRowTwo.getCell(0), "测试", 3, true, 3,"BFBFBF", STVerticalJc.CENTER, STJc.CENTER);
		// tableOneRowTwo.getCell(0).

		// os.close();
		// table = tables.get(3);
		// ttbl = table.getCTTbl();
		// tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() :
		// ttbl.getTblPr();
		// rows = table.getRows();
		// for (XWPFTableRow row : rows) {
		// //获取行对应的单元格
		// cells = row.getTableCells();
		// for (XWPFTableCell cell : cells) {
		// System.out.println(cell.getText());;
		// }
		// }
		// tableOneRowTwo = table.createRow();
		// tableOneRowTwo.getCell(0).setText("第2行第1列");
		// for (XWPFTable table : tables) {
		// //表格属性
		// // CTTblPr pr = table.getCTTbl().getTblPr();
		// CTTbl ttbl = table.getCTTbl();
		// CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() :
		// ttbl.getTblPr();
		// //获取表格对应的行
		// rows = table.getRows();
		// for (XWPFTableRow row : rows) {
		// //获取行对应的单元格
		// cells = row.getTableCells();
		// for (XWPFTableCell cell : cells) {
		// System.out.println(cell.getText());;
		// }
		// }
		// XWPFTableRow tableOneRowTwo = table.createRow();
		// tableOneRowTwo.getCell(0).setText("第2行第1列");
		// }
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
        CTTblWidth cTblWidth = ctPr.addNewTcW();  
//        cTblWidth.setW(BigInteger.valueOf(width));  
//        cTblWidth.setType(STTblWidth.Enum.forString("dxa"));  
        if (isShd) {  
            if (shdValue > 0 && shdValue <= 38) {  
                ctshd.setVal(STShd.Enum.forInt(shdValue));  
            }  
            if (shdColor != null) {  
                ctshd.setColor(shdColor);  
//                ctshd.set
            }  
        }  
        ctPr.addNewVAlign().setVal(jc);  
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(stJc); 
        XWPFParagraph cellP=cell.getParagraphs().get(0);
		XWPFRun cellR = cellP.createRun();
		cellR.setFontSize(11);
		cellR.setText(text);
//        cell.setText(text);  
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
//		CTTcPr cellPr = cttc.addNewTcPr();
//		cellPr.addNewTcW().setW(BigInteger.valueOf(width));
		// cell.setColor(bgcolor);
		CTTcPr ctPr = cttc.addNewTcPr();
//		CTFonts fonts = ctPr.isSetRFonts() ? ctPr.getRFonts() : ctPr.addNewRFonts();
//		fonts.setAscii("宋体");
//		fonts.setEastAsia("宋体");
//		fonts.setHAnsi("宋体");
		
//		CTShd ctshd = ctPr.addNewShd();
//		ctshd.setFill(bgcolor);
		
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
//		ctPr.set
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
//		cttc.set
        XWPFParagraph cellP=cell.getParagraphs().get(0);
		XWPFRun cellR = cellP.createRun();
		cellR.setFontSize(10);
		cellR.setText(text);
//		cell.setText(text);
	}

	public void addNewPage(XWPFDocument document,BreakType breakType){
		XWPFParagraph xp = document.createParagraph();
		xp.createRun().addBreak(breakType);
	}
	
	/**
	 * @Description:按位置得到单元格(考虑跨列合并情况)
	 */
	public XWPFTableCell getCellSizeWithMergeNum(XWPFTableRow row, int position) {
		List<XWPFTableCell> rowCellList = row.getTableCells();
		int k = 0;
		for (int i = 0, len = rowCellList.size(); i < len; i++) {
			CTTc ctTc = rowCellList.get(i).getCTTc();
			if (ctTc.isSetTcPr()) {
				CTTcPr tcPr = ctTc.getTcPr();
				if (tcPr.isSetGridSpan()) {
					CTDecimalNumber gridSpan = tcPr.getGridSpan();
					k += gridSpan.getVal().intValue() - 1;
				}
			}
			if (k >= position) {
				return rowCellList.get(i);
			}
			k++;
		}
		if (position < rowCellList.size()) {
			return rowCellList.get(position);
		}
		return null;
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
			XwpfTest xwpf = new XwpfTest();
			xwpf.testReadByDoc();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}