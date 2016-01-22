package com.cm.oe.budget.gen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JFrame;
import javax.swing.ProgressMonitorInputStream;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.record.PageBreakRecord.Break;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFFormulaEvaluator.RowFlushedException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.w3c.dom.css.ElementCSSInlineStyle;

import com.cm.oe.test.ReadExcel;
import com.cm.oe.test.ReadExcelTable;

public class BudgetWriter1 {
	private String path1 = "testfiles/ysb_final.xls";
	private String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private String tablePath = "testfiles/tables.xls";
	private String excelPath = "testfiles/testall.xls";
	private String output = "testfiles/";
	private BudgetReader1 ub = new BudgetReader1(path1, path2);
	private BudgetReader2 ub2 = new BudgetReader2(path1, path2);
	private BudgetReader3 ub3 = new BudgetReader3(path1, path2);
	private ReadExcelTable ret = new ReadExcelTable();
	private ReadExcel re = new ReadExcel();

	public BudgetWriter1(String path1, String path2,  String tablePath, String excelPath,
			String output) {
		this.path1 = path1;
		this.path2 = path2;
		this.tablePath = tablePath;
		this.excelPath = excelPath;
		this.output = output;
	}
	public static void main(String[] args) throws Exception {
		 String path1 = "testfiles/ysb_final.xls";
		 String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
		 String tablePath = "testfiles/tables.xls";
		 String excelPath = "C:\\Users\\王宁\\Desktop\\新建文件夹\\宏基站-室外站-D频段-上海贝尔-24111.xls";
		 String output = "C:\\Users\\王宁\\Desktop\\新建文件夹\\";
		BudgetWriter1 bw = new BudgetWriter1(path1, path2,  tablePath, excelPath, output);
		bw.testReadByDoc();
	}
	public void testReadByDoc() throws Exception {
		Map<Integer, List<String>> excelmap = new HashMap<Integer, List<String>>();
		Map<Integer, List<String>> BBUtablemap = ret.readBBUinExcel(tablePath, excelPath);
		Map<Integer, List<String>> RRUtablemap = ret.readRRUinExcel(tablePath, excelPath);
		Map<Integer, List<String>> Antennatablemap = ret.readAntennaIntables(tablePath, excelPath);
		Map<Integer, String> AZHoLa = ret.anzhuang(tablePath, excelPath);
		Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
		Map<String, List<String>> data_b3 = ub.getB3(path1);
		Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path1, path2);
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		Sheet sheet = wb.getSheetAt(0);

		int rowNums = re.rowNumber(wb);
		FileOutputStream fos = null;
		Row r = null;
		String city = null;

		for (int i = 0; i < rowNums; i++) {
			List<String> line = new ArrayList<String>();
			r = sheet.getRow(i);
			String zmzh = re.getExcelvalues(r).get(1).trim() + re.getExcelvalues(r).get(2).trim();
			String dizhi = re.getExcelvalues(r).get(0).trim();
			line.add(dizhi);
			line.add(zmzh);
			for(int j =3;j<re.getExcelvalues(r).size();j++){
				line.add(re.getExcelvalues(r).get(j).trim());
			}
			excelmap.put(i, line);
		}
		//System.out.println(excelmap);
		//System.out.println(RRUtablemap);
		for (int i = 1; i < excelmap.size(); i++) {
			File file = new File(excelPath);
			InputStream is = null ;
			String[] strs = file.getName().split("-");
			String  name =strs[0];
			String neiwai = strs[1];
			String pinduan = strs[2];
			String shebei = strs[3];
			if(name.trim().equals("宏基站")){
				if(neiwai.trim().equals("室内站")){
					is = new FileInputStream("testfiles/宏基站室内.docx");
				}else if(neiwai.trim().equals("室外站")){
					is = new FileInputStream("testfiles/宏基站室外.docx");
				}
				
			}else if(name.trim().equals("小基站")){
				if(neiwai.trim().equals("室内站")){
					is = new FileInputStream("testfiles/小基站室内.docx");
				}else if(neiwai.trim().equals("室外站")){
					is = new FileInputStream("testfiles/小基站室外.docx");
				}
			}else if(name.trim().equals("拉远站")){
				if(neiwai.trim().equals("室内站")){
					is = new FileInputStream("testfiles/拉远站室内.docx");
				}else if(neiwai.trim().equals("室外站")){
					is = new FileInputStream("testfiles/拉远站室外.docx");
				}
			}else if(name.trim().equals("信源站")){
				if(neiwai.trim().equals("室内站")){
					is = new FileInputStream("testfiles/信源站室内.docx");
				}else if(neiwai.trim().equals("室外站")){
					is = new FileInputStream("testfiles/信源站室外.docx");
				}
			}
			XWPFDocument doc = new XWPFDocument(is);
			List<XWPFTable> tables = doc.getTables();
			if(name.trim().equals("信源站")){
				XWPFTable table0 = tables.get(0);
				setCellText(table0.getRow(0).getCell(1),excelmap.get(i).get(26).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(1).getCell(1),pinduan.toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(2).getCell(1),name.toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(3).getCell(1),excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(4).getCell(1),excelmap.get(i).get(28).toString().trim(), "FFFFFF", 21);
			}else if(name.trim().equals("小基站")||name.trim().equals("拉远站")||name.trim().equals("宏基站")){
				XWPFTable table0 = tables.get(0);
				//TODO excelmap里的顺序需要重新改过
				setCellText(table0.getRow(0).getCell(1), excelmap.get(i).get(26).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(1).getCell(1), pinduan.toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(2).getCell(1), name.toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(31).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(32).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(0).getCell(3), shebei.toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(28).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(29).toString().trim(), "FFFFFF", 21);
				setCellText(table0.getRow(4).getCell(3), excelmap.get(i).get(30).toString().trim(), "FFFFFF", 21);
			}
			System.out.println(excelmap.get(0).get(18).toString());
			System.out.println(excelmap.get(0).get(18).toString().trim().equals("BBU型号"));
			if (excelmap.get(0).get(19).toString().trim().equals("BBU型号")) {
				XWPFTable tableBBU = tables.get(2);
				XWPFTableRow tBBURow = tableBBU.createRow();
				tBBURow.setHeight(11);
				//System.out.println(BBUtablemap);
				for (int j = 0; j < BBUtablemap.get(i).size(); j++) {
					setCellText(tBBURow.getCell(j), BBUtablemap.get(i).get(j), "FFFFFF", 21);
				}
			}
			if (excelmap.get(0).get(21).toString().trim().equals("RRU型号")) {
				XWPFTable tableRRU = tables.get(3);
				XWPFTableRow row = tableRRU.getRow(0);
				//System.out.println(RRUtablemap.get(4));
				if (RRUtablemap.get(i).get(8).toString().equals("工作频带宽度 ")) {
					mergeCellsHorizontal(tableRRU, 0, 5, 7);
					XWPFTableRow tRRURow = tableRRU.createRow();
					mergeCellsHorizontal(tableRRU, 1, 5, 7);
					tRRURow.setHeight(11);
					setCellText(tRRURow.getCell(0), RRUtablemap.get(i).get(0), "FFFFFF", 21);
					setCellText(tRRURow.getCell(1), RRUtablemap.get(i).get(1), "FFFFFF", 21);
					setCellText(tRRURow.getCell(2), RRUtablemap.get(i).get(2), "FFFFFF", 21);
					setCellText(tRRURow.getCell(3), RRUtablemap.get(i).get(3), "FFFFFF", 21);
					setCellText(tRRURow.getCell(4), RRUtablemap.get(i).get(4), "FFFFFF", 21);
					setCellText(tRRURow.getCell(5), RRUtablemap.get(i).get(5), "FFFFFF", 21);
					setCellText(tRRURow.getCell(8), RRUtablemap.get(i).get(6), "FFFFFF", 21);
					setCellText(tRRURow.getCell(9), RRUtablemap.get(i).get(7), "FFFFFF", 21);
				}
				//System.out.println(RRUtablemap);
				if (RRUtablemap.get(i).get(8).toString().equals("功耗")) {
					mergeCellsHorizontal(tableRRU, 0, 5, 6);
					XWPFTableCell cell = row.getCell(5);
					cell.removeParagraph(0);
					cell.setText("供电方式");
					XWPFTableRow tRRURow = tableRRU.createRow();
					mergeCellsHorizontal(tableRRU, 1, 5, 6);
					tRRURow.setHeight(11);

					setCellText(tRRURow.getCell(0), RRUtablemap.get(i).get(0), "FFFFFF", 21);
					setCellText(tRRURow.getCell(1), RRUtablemap.get(i).get(1), "FFFFFF", 21);
					setCellText(tRRURow.getCell(2), RRUtablemap.get(i).get(2), "FFFFFF", 21);
					setCellText(tRRURow.getCell(3), RRUtablemap.get(i).get(3), "FFFFFF", 21);
					setCellText(tRRURow.getCell(4), RRUtablemap.get(i).get(4), "FFFFFF", 21);
					setCellText(tRRURow.getCell(5), RRUtablemap.get(i).get(9), "FFFFFF", 21);
					setCellText(tRRURow.getCell(7), RRUtablemap.get(i).get(5), "FFFFFF", 21);
					setCellText(tRRURow.getCell(8), RRUtablemap.get(i).get(6), "FFFFFF", 21);
					setCellText(tRRURow.getCell(9), RRUtablemap.get(i).get(7), "FFFFFF", 21);

				}
			}	
				if (excelmap.get(0).get(22).toString().trim().equals("天线型号")) {
					XWPFTable tableAnn = tables.get(4);
					XWPFTableRow tAnnRow = tableAnn.createRow();
					tAnnRow.setHeight(11);
					for (int j = 0; j < Antennatablemap.get(i).size(); j++) {
						setCellText(tAnnRow.getCell(j), Antennatablemap.get(i).get(j), "FFFFFF", 21);
					}
				}
				
				XWPFTable table = tables.get(1);
				CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
				genBorders(borders);
				for (String key : datas_map.keySet()) {
					XWPFTableRow tableOneRowTwo = table.createRow();
					tableOneRowTwo.setHeight(11);
					setCellText(tableOneRowTwo.getCell(0), key, "FFFFFF", 21);
					setCellText(tableOneRowTwo.getCell(1), datas_map.get(key).get(0), "FFFFFF", 21);
					setCellText(tableOneRowTwo.getCell(2), datas_map.get(key).get(1), "FFFFFF", 21);
				}

				String zh = ub2.getZhFrom4GYsb();
				List<String> datas = ub2.readExcel(zh);
				if(name.toString().equals("信源站")||name.toString().equals("拉远站")){
					table = tables.get(5);
				}else{
					table = tables.get(6);
				}
				borders = table.getCTTbl().getTblPr().addNewTblBorders();
				genBorders(borders);
				XWPFTableRow tableOneRowTwo = table.createRow();
				tableOneRowTwo.setHeight(11);
				int k = 0;
				for (String values : datas) {
					setCellText(tableOneRowTwo.getCell(k), values, "FFFFFF", 21);
					k++;
				}

				datas = ub3.readExcel(zh);
				if(name.toString().trim().equals("信源站")||name.toString().trim().equals("拉远站")){
					table = tables.get(6);
				}else {
					table = tables.get(7);
				}
				borders = table.getCTTbl().getTblPr().addNewTblBorders();
				genBorders(borders);
				tableOneRowTwo = table.createRow();
				tableOneRowTwo.setHeight(11);
				k = 0;
				for (String values : datas) {
					setCellText(tableOneRowTwo.getCell(k), values, "FFFFFF", 21);
					k++;
				}

				
				//OPCPackage pack = POIXMLDocument.openPackage(template);
				//XWPFDocument doc = new XWPFDocument(pack);
				List<XWPFParagraph> paragraphs = doc.getParagraphs();
				//System.out.println(paragraphs.size());
				for (XWPFParagraph tmp : paragraphs) {
					//System.out.println(tmp.getParagraphText());
					List<XWPFRun> runs = tmp.getRuns();
					for (XWPFRun aa : runs) {
						//System.out.println("XWPFRun-Text:" + aa.getText(0));
						if ("city".equals(aa.getText(0))) {
							aa.setText(excelmap.get(i).get(0), 0);
						}
						if("option2".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(1),0);
						}
						if("option3".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(2),0);
						}
						if("option4".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(3),0);
						}
						if("option5".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(4),0);
						}
						if("option6".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(5),0);
						}
						if("option7".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(6),0);
						}
						if("option8".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(7),0);
						}
						if("option9".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(8),0);
						}
						if("option10".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(9),0);
						}
						if("option11".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(10),0);
						}
						if("option12".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(11),0);
						}
						if("option13".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(12),0);
						}
						if("option14".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(13),0);
						}
						if("option15".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(14),0);
						}
						if("option16".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(19),0);
						}
						if("option17".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(21),0);
						}
						if("option18".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(22),0);
						}
						if("option19".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(21),0);
						}
						if("hongjizhan".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(13),0);
						}
						if("xiaojizhan".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(14),0);
						}
						if("layuanzhan".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(15),0);
						}
						if("xinyuanzhan".equals(aa.getText(0))){
							aa.setText(excelmap.get(i).get(16),0);
						}
						//System.out.println(excelmap.get(i).get(25).toString().trim().equals("挂墙机框"));
						//System.out.println(ret.anzhuangHongandLa(tablePath, excelPath));
						//System.out.println(AZHoLa.get(0));
						if("anzhuangfangshi".equals(aa.getText(0))){
							aa.setText(AZHoLa.get(i),0);
						}
					}
				name = excelmap.get(i).get(0);
				fos = new FileOutputStream(output + city + ".doc");
				doc.write(fos);
				fos.flush();
				fos.close();
				is.close();
			}
			wb.close();
			fise.close();
		}
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

	public void setCellText(XWPFTableCell cell, String text, String bgcolor, int width) {
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
		XWPFParagraph cellP = cell.getParagraphs().get(0);
		XWPFRun cellR = cellP.createRun();
		cellR.setFontSize(10);
		cellR.setText(text);
	}


	private void close(InputStream is) {
		if (is != null) {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCell) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one, are set with
				// CONTINUE
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}
}