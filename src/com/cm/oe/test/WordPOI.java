package com.cm.oe.test;


import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

public class WordPOI {
	public static void main(String[] args) throws Exception {
		WordPOI t = new WordPOI();
		t.testReplaceValue("testfiles/sys_s_tmp_8.docx");
	}

	public void testReplaceValue(String fileName) throws Exception {
		XWPFDocument xdoc = new XWPFDocument(
				POIXMLDocument.openPackage(fileName));
		Map<String, String> paramMap = new HashMap<String, String>();
		paramMap.put("${vdate}", "2014-11-21");
		paramMap.put("${version}", "V1.1");
		paramMap.put("${versiondesc}", "仅供参考,From小瓜的博客");
		paramMap.put("${v_author}", "小瓜");
		paramMap.put("${name_1}", "测试POI替换Word 2007值");
		paramMap.put("${s_name}", "小瓜");
		paramMap.put("${no}", "" + System.currentTimeMillis());
		paramMap.put("${class_name}", "保密");
		paramMap.put("${t_name}", "瓜哥");
		paramMap.put("${t_level}", "讲师");
		paramMap.put("${major_name}", "计科");
		paramMap.put("${t_desc}", "满架蔷薇一院香");
		paramMap.put("${t_desc_date_1}", "2014-10-1");
		paramMap.put("${t_desc_date_2}", "2014-10-1");
		paramMap.put("${t_desc_2}", "风过栏杆水不波");
		paramMap.put("${t_desc_date_3}", "2014-10-2");
		paramMap.put("${t_desc_date_4}", "2014-10-2");
		paramMap.put("${t_desc_3}", "此去经年,应是良辰好景虚设");
		paramMap.put("${t_desc_date_5}", "2014-10-3");
		paramMap.put("${t_desc_date_6}", "2014-10-3");
		paramMap.put("${school_name}", "大学中庸");
		paramMap.put("${b_j}", "2015");
		paramMap.put("${b_name}", "无线路由器");
		paramMap.put("${b_marjor}", "接天莲叶无穷碧");
		paramMap.put("${b_class}", "无才");
		paramMap.put("${b_date}", "二〇一四年十一月二十一日");
		paramMap.put("${c_name}", "莫说相公痴");
		paramMap.put("${c_add}", "般若波罗蜜多心经");
		paramMap.put("${c_mm}", "如梦亦如幻");
		paramMap.put("${c_mobile}", "因作如是观");
		paramMap.put("${cn_name}", "订单表");
		paramMap.put("${en_name}", "oracle_order_table");
		paramMap.put("${table_desc}", "订单表测试");
		List<List<String>> cellList=generateTestData(8);
		replaceParagraphValue(xdoc, paramMap);
		replaceTableValueNoraml(xdoc, paramMap);
		replaceLastTableValue(xdoc, paramMap, cellList);
		saveDocument(xdoc, "testfiles/sys_" +System.currentTimeMillis()+ ".docx");
	}

	public List<List<String>> generateTestData(int num) {
		List<List<String>> resultList = new ArrayList<List<String>>();
		for (int i = 1; i <= num; i++) {
			List<String> list = new ArrayList<String>();
			list.add("测试_" + i);
			list.add("测试2_" + i);
			list.add("varchar2(60)");
			list.add("无");
			list.add("String");
			list.add("否");
			list.add("是");
			resultList.add(list);
		}
		return resultList;
	}
	
	public void replaceParagraphValue(XWPFDocument xdoc,
			Map<String, String> paramMap) throws Exception {
		// 替换段落中的指定文字
		List<XWPFParagraph> paragraphList = xdoc.getParagraphs();
		if (paragraphList != null && paragraphList.size() > 0) {
			for (XWPFParagraph paragraph : paragraphList) {
				List<XWPFRun> runs = paragraph.getRuns();
				if (runs == null || runs.size() == 0) {
					continue;
				}
				for (int i = 0, len = runs.size(); i < len; i++) {
					String oldStr = runs.get(i).getText(0);
					if(oldStr==null){
						continue;
					}
					for (Entry<String, String> e : paramMap.entrySet()) {
						oldStr = oldStr.replace(e.getKey(), e.getValue());
					}
					if (oldStr != null) {
						runs.get(i).setText(new String(oldStr), 0);
					}
				}
			}
		}
	}

	// 不替换最后一个表格
	public void replaceTableValueNoraml(XWPFDocument xdoc,
			Map<String, String> paramMap) throws Exception {
		// 替换表格中的指定文字
		List<XWPFTable> tableList = xdoc.getTables();
		// 不替换最后一个表格
		for (int i = 0, len = tableList.size() - 1; i < len; i++) {
			XWPFTable table = tableList.get(i);
			for (int j = 0, rcount = table.getNumberOfRows(); j < rcount; j++) {
				XWPFTableRow row = table.getRow(j);
				List<XWPFTableCell> cells = row.getTableCells();
				if (cells == null || cells.size() == 0) {
					continue;
				}
				for (XWPFTableCell cell : cells) {
					List<XWPFParagraph> cellPList = cell.getParagraphs();
					replaceTableCellParagraphValue(xdoc, cellPList, paramMap);
				}
			}
		}
	}

	public void replaceTableCellParagraphValue(XWPFDocument xdoc,
			List<XWPFParagraph> cellPList, Map<String, String> paramMap)
			throws Exception {
		// 替换段落中的指定文字
		if (cellPList != null && cellPList.size() > 0) {
			for (XWPFParagraph paragraph : cellPList) {
				List<XWPFRun> runs = paragraph.getRuns();
				if (runs == null || runs.size() == 0) {
					continue;
				}
				for (int i = 0, len = runs.size(); i < len; i++) {
					String oldStr = runs.get(i).getText(0);
					if(oldStr==null){
						continue;
					}
					for (Entry<String, String> e : paramMap.entrySet()) {
						oldStr = oldStr.replace(e.getKey(), e.getValue());
					}
					if (oldStr != null) {
						runs.get(i).setText(new String(oldStr), 0);
					}
				}
			}
		}
	}

	//替换最后一个表格值 
	public void replaceLastTableValue(XWPFDocument xdoc,
			Map<String, String> paramMap, List<List<String>> resultList)
			throws Exception {
		List<XWPFTable> tableList = xdoc.getTables();
		if (tableList == null || tableList.size() == 0) {
			return;
		}
		XWPFTable table = tableList.get(tableList.size() - 1);
		XWPFTableRow row = null;
		List<XWPFTableCell> cells = null;
		//替换模版行前面的数据
		for (int j = 0, rcount = table.getNumberOfRows() - 1; j < rcount; j++) {
			row = table.getRow(j);
			cells = row.getTableCells();
			if (cells == null || cells.size() == 0) {
				continue;
			}
			for (int k = 0, len = cells.size(); k < len; k++) {
				List<XWPFParagraph> cellPList = cells.get(k).getParagraphs();
				replaceTableCellParagraphValue(xdoc, cellPList, paramMap);
			}
		}
		int lastRowSize=table.getNumberOfRows() - 1;
		XWPFTableRow lastRow = table.getRow(lastRowSize);
		List<XWPFTableCell> tmpCells = lastRow.getTableCells();
		if (tmpCells.size() != resultList.get(0).size()) {
			return;
		}
		XWPFTableCell tmpCell = null;
		for (int i = 0, len = resultList.size(); i < len; i++) {
			row = table.createRow();
			row.setHeight(lastRow.getHeight());
			List<String> list = resultList.get(i);
			cells = row.getTableCells();
			// 插入的行会填充与表格第一行相同的列数
			for (int k = 0, klen = cells.size(); k < klen; k++) {
				tmpCell = tmpCells.get(k);
				XWPFTableCell cell = cells.get(k);
				setCellText(tmpCell, cell, list.get(k));
			}
			// 继续写剩余的列
			for (int j = cells.size(), jlen = list.size(); j < jlen; j++) {
				tmpCell = tmpCells.get(j);
				XWPFTableCell cell = row.addNewTableCell();
				setCellText(tmpCell, cell, list.get(j));
			}
		}
		//删除模版行
		table.removeRow(lastRowSize);
	}

	//非完全复制样式(只复制简单的样式)
	public void setCellText(XWPFTableCell tmpCell, XWPFTableCell cell,String text) throws Exception {
		CTTc cttc2 = tmpCell.getCTTc();
		CTTcPr ctPr2 = cttc2.getTcPr();

		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		cell.setColor(tmpCell.getColor());
		cell.setVerticalAlignment(tmpCell.getVerticalAlignment());
		if (ctPr2.getTcW() != null) {
			ctPr.addNewTcW().setW(ctPr2.getTcW().getW());
		}
		if (ctPr2.getVAlign() != null) {
			ctPr.addNewVAlign().setVal(ctPr2.getVAlign().getVal());
		}
		if (cttc2.getPList().size() > 0) {
			CTP ctp = cttc2.getPList().get(0);
			if (ctp.getPPr() != null) {
				if (ctp.getPPr().getJc() != null) {
					cttc.getPList().get(0).addNewPPr().addNewJc()
							.setVal(ctp.getPPr().getJc().getVal());
				}
			}
		}

		if (ctPr2.getTcBorders() != null) {
			ctPr.setTcBorders(ctPr2.getTcBorders());
		}

		XWPFParagraph tmpP = tmpCell.getParagraphs().get(0);
		XWPFParagraph cellP = cell.getParagraphs().get(0);
		XWPFRun tmpR = null;
		if (tmpP.getRuns() != null && tmpP.getRuns().size() > 0) {
			tmpR = tmpP.getRuns().get(0);
		}
		XWPFRun cellR = cellP.createRun();
		cellR.setText(text);
		// 复制字体信息
		if (tmpR != null) {
			cellR.setBold(tmpR.isBold());
			cellR.setItalic(tmpR.isItalic());
			cellR.setStrike(tmpR.isStrike());
			cellR.setUnderline(tmpR.getUnderline());
			cellR.setColor(tmpR.getColor());
			cellR.setTextPosition(tmpR.getTextPosition());
			if (tmpR.getFontSize() != -1) {
				cellR.setFontSize(tmpR.getFontSize());
			}
			if (tmpR.getFontFamily() != null) {
				cellR.setFontFamily(tmpR.getFontFamily());
			}
			if (tmpR.getCTR() != null) {
				if (tmpR.getCTR().isSetRPr()) {
					CTRPr tmpRPr = tmpR.getCTR().getRPr();
					if (tmpRPr.isSetRFonts()) {
						CTFonts tmpFonts = tmpRPr.getRFonts();
						CTRPr cellRPr = cellR.getCTR().isSetRPr() ? cellR
								.getCTR().getRPr() : cellR.getCTR().addNewRPr();
						CTFonts cellFonts = cellRPr.isSetRFonts() ? cellRPr
								.getRFonts() : cellRPr.addNewRFonts();
						cellFonts.setAscii(tmpFonts.getAscii());
						cellFonts.setAsciiTheme(tmpFonts.getAsciiTheme());
						cellFonts.setCs(tmpFonts.getCs());
						cellFonts.setCstheme(tmpFonts.getCstheme());
						cellFonts.setEastAsia(tmpFonts.getEastAsia());
						cellFonts.setEastAsiaTheme(tmpFonts.getEastAsiaTheme());
						cellFonts.setHAnsi(tmpFonts.getHAnsi());
						cellFonts.setHAnsiTheme(tmpFonts.getHAnsiTheme());
					}
				}
			}
		}
		// 复制段落信息
		cellP.setAlignment(tmpP.getAlignment());
		cellP.setVerticalAlignment(tmpP.getVerticalAlignment());
		cellP.setBorderBetween(tmpP.getBorderBetween());
		cellP.setBorderBottom(tmpP.getBorderBottom());
		cellP.setBorderLeft(tmpP.getBorderLeft());
		cellP.setBorderRight(tmpP.getBorderRight());
		cellP.setBorderTop(tmpP.getBorderTop());
		cellP.setPageBreak(tmpP.isPageBreak());
		if (tmpP.getCTP() != null) {
			if (tmpP.getCTP().getPPr() != null) {
				CTPPr tmpPPr = tmpP.getCTP().getPPr();
				CTPPr cellPPr = cellP.getCTP().getPPr() != null ? cellP
						.getCTP().getPPr() : cellP.getCTP().addNewPPr();
				// 复制段落间距信息
				CTSpacing tmpSpacing = tmpPPr.getSpacing();
				if (tmpSpacing != null) {
					CTSpacing cellSpacing = cellPPr.getSpacing() != null ? cellPPr
							.getSpacing() : cellPPr.addNewSpacing();
					if (tmpSpacing.getAfter() != null) {
						cellSpacing.setAfter(tmpSpacing.getAfter());
					}
					if (tmpSpacing.getAfterAutospacing() != null) {
						cellSpacing.setAfterAutospacing(tmpSpacing
								.getAfterAutospacing());
					}
					if (tmpSpacing.getAfterLines() != null) {
						cellSpacing.setAfterLines(tmpSpacing.getAfterLines());
					}
					if (tmpSpacing.getBefore() != null) {
						cellSpacing.setBefore(tmpSpacing.getBefore());
					}
					if (tmpSpacing.getBeforeAutospacing() != null) {
						cellSpacing.setBeforeAutospacing(tmpSpacing
								.getBeforeAutospacing());
					}
					if (tmpSpacing.getBeforeLines() != null) {
						cellSpacing.setBeforeLines(tmpSpacing.getBeforeLines());
					}
					if (tmpSpacing.getLine() != null) {
						cellSpacing.setLine(tmpSpacing.getLine());
					}
					if (tmpSpacing.getLineRule() != null) {
						cellSpacing.setLineRule(tmpSpacing.getLineRule());
					}
				}
				// 复制段落缩进信息
				CTInd tmpInd = tmpPPr.getInd();
				if (tmpInd != null) {
					CTInd cellInd = cellPPr.getInd() != null ? cellPPr.getInd()
							: cellPPr.addNewInd();
					if (tmpInd.getFirstLine() != null) {
						cellInd.setFirstLine(tmpInd.getFirstLine());
					}
					if (tmpInd.getFirstLineChars() != null) {
						cellInd.setFirstLineChars(tmpInd.getFirstLineChars());
					}
					if (tmpInd.getHanging() != null) {
						cellInd.setHanging(tmpInd.getHanging());
					}
					if (tmpInd.getHangingChars() != null) {
						cellInd.setHangingChars(tmpInd.getHangingChars());
					}
					if (tmpInd.getLeft() != null) {
						cellInd.setLeft(tmpInd.getLeft());
					}
					if (tmpInd.getLeftChars() != null) {
						cellInd.setLeftChars(tmpInd.getLeftChars());
					}
					if (tmpInd.getRight() != null) {
						cellInd.setRight(tmpInd.getRight());
					}
					if (tmpInd.getRightChars() != null) {
						cellInd.setRightChars(tmpInd.getRightChars());
					}
				}
			}
		}
	}

	public void saveDocument(XWPFDocument document, String savePath)
			throws Exception {
		FileOutputStream fos = new FileOutputStream(savePath);
		document.write(fos);
		fos.close();
	}
}
