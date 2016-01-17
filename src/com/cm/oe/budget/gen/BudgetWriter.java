package com.cm.oe.budget.gen;

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


public class BudgetWriter {
	private String template = "";
	private String path1 = "testfiles/ysb_final.xls";
	private String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private String output = "";
	private BudgetReader1 ub = new BudgetReader1(path1, path2);
	private BudgetReader2 ub2 = new BudgetReader2(path1, path2);
	private BudgetReader3 ub3 = new BudgetReader3(path1, path2);

	public BudgetWriter(String path1, String path2, String template, String output){
		this.template = template;
		this.path1 = path1;
		this.path2 = path2;
		this.output = output;
	}
	
	public void testReadByDoc() throws Exception {
		InputStream is = new FileInputStream(template);
		OutputStream os = new FileOutputStream(output);
		XWPFDocument doc = new XWPFDocument(is);
		Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
		Map<String, List<String>> data_b3 = ub.getB3(path1);
		Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path1, path2);
		List<XWPFTable> tables = doc.getTables();

		XWPFTable table = tables.get(2);
		CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
		genBorders(borders);
		for(String key:datas_map.keySet()){
			XWPFTableRow tableOneRowTwo = table.createRow();
			tableOneRowTwo.setHeight(11);
			setCellText(tableOneRowTwo.getCell(0), key, "FFFFFF", 21);
			setCellText(tableOneRowTwo.getCell(1), datas_map.get(key).get(0), "FFFFFF", 21);
			setCellText(tableOneRowTwo.getCell(2), datas_map.get(key).get(1), "FFFFFF", 21);
		}

		String zh = ub2.getZhFrom4GYsb();
		List<String> datas = ub2.readExcel(zh);
		table = tables.get(7);
		borders = table.getCTTbl().getTblPr().addNewTblBorders();
		genBorders(borders);
		XWPFTableRow tableOneRowTwo = table.createRow();
		tableOneRowTwo.setHeight(11);
		int i = 0;
		for(String values:datas){
			setCellText(tableOneRowTwo.getCell(i), values, "FFFFFF", 21);
			i++;
		}
		
		datas = ub3.readExcel(zh);
		table = tables.get(8);
		borders = table.getCTTbl().getTblPr().addNewTblBorders();
		genBorders(borders);
		tableOneRowTwo = table.createRow();
		tableOneRowTwo.setHeight(11);
		i = 0;
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
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        XWPFParagraph cellP=cell.getParagraphs().get(0);
		XWPFRun cellR = cellP.createRun();
		cellR.setFontSize(10);
		cellR.setText(text);
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
		String filePath = "testfiles/templete_all.docx";
		String path1 = "testfiles/ysb_final.xls";
		String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
		String output = "testfiles/templete_all2.docx";
		try {
			BudgetWriter xwpf = new BudgetWriter(path1, path2, filePath, output);
			xwpf.testReadByDoc();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}