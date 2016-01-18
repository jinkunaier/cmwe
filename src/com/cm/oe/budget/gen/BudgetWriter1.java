package com.cm.oe.budget.gen;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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

import com.cm.oe.test.ReadExcel;
import com.cm.oe.test.ReadExcelTable;

public class BudgetWriter1 {
	private String template = "";
	private String path1 = "testfiles/ysb_final.xls";
	private String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private String tablePath = "testfiles/tables.xls";
	private String excelPath = "testfiles/test.xls";
	private String output = "testfiles/";
	private BudgetReader1 ub = new BudgetReader1(path1, path2);
	private BudgetReader2 ub2 = new BudgetReader2(path1, path2);
	private BudgetReader3 ub3 = new BudgetReader3(path1, path2);
	private ReadExcelTable ret = new ReadExcelTable();
	private ReadExcel re = new ReadExcel();

	public BudgetWriter1(String path1, String path2, String template, String tablePath, String excelPath,
			String output) {
		this.template = template;
		this.path1 = path1;
		this.path2 = path2;
		this.tablePath = tablePath;
		this.excelPath = excelPath;
		this.output = output;
	}

	public void testReadByDoc() throws Exception {
		Map<Integer, List<String>> excelmap = new HashMap<Integer, List<String>>();
		Map<Integer, List<String>> BBUtablemap = ret.readBBUinExcel(tablePath, excelPath);
		Map<Integer, List<String>> RRUtablemap = ret.readRRUinExcel(tablePath, excelPath);
		Map<Integer, List<String>> Antennatablemap = ret.readAntennaIntables(tablePath, excelPath);
		Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
		Map<String, List<String>> data_b3 = ub.getB3(path1);
		Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path1, path2);

		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		Sheet sheet = wb.getSheetAt(0);

		int rowNums = re.rowNumber(wb);
		FileOutputStream fos = null;
		Row r = null;
		String name = null;

		for (int i = 0; i < rowNums; i++) {
			r = sheet.getRow(i);
			excelmap.put(i, re.getExcelvalues(r));
		}

		for (int i = 1; i < excelmap.size(); i++) {
			InputStream is = new FileInputStream(template);
			XWPFDocument doc = new XWPFDocument(is);
			List<XWPFTable> tables = doc.getTables();

			if (excelmap.get(0).get(6).toString().equals("BBU")) {
				XWPFTable tableBBU = tables.get(3);
				XWPFTableRow tBBURow = tableBBU.createRow();
				tBBURow.setHeight(11);
				//System.out.println(BBUtablemap);
				for (int j = 0; j < BBUtablemap.get(i).size(); j++) {
					setCellText(tBBURow.getCell(j), BBUtablemap.get(i).get(j), "FFFFFF", 21);
				}
			}
			if (excelmap.get(0).get(7).toString().equals("RRU")) {
				XWPFTable tableRRU = tables.get(4);
				XWPFTableRow row = tableRRU.getRow(0);

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
				System.out.println(RRUtablemap);
				if (RRUtablemap.get(i).get(8).toString().equals("功耗")) {
					mergeCellsHorizontal(tableRRU, 0, 5, 6);
					XWPFTableCell cell = row.getCell(5);
					cell.removeParagraph(0);
					cell.setText("供电方式");
					XWPFTableRow tRRURow = tableRRU.createRow();
					mergeCellsHorizontal(tableRRU, 1, 5, 6);
					tRRURow.setHeight(11);

					//System.out.println(RRUtablemap.get(i).get(5));
					System.out.println(RRUtablemap.get(i));
					System.out.println(RRUtablemap.get(i).get(2));
					System.out.println(RRUtablemap.get(i).get(6));
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
				if (excelmap.get(0).get(8).toString().equals("Antenna")) {
					XWPFTable tableAnn = tables.get(5);
					XWPFTableRow tAnnRow = tableAnn.createRow();
					tAnnRow.setHeight(11);
					for (int j = 0; j < Antennatablemap.get(i).size(); j++) {
						setCellText(tAnnRow.getCell(j), Antennatablemap.get(i).get(j), "FFFFFF", 21);
					}
				}
				XWPFTable table = tables.get(2);
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
				table = tables.get(7);
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
				table = tables.get(8);
				borders = table.getCTTbl().getTblPr().addNewTblBorders();
				genBorders(borders);
				tableOneRowTwo = table.createRow();
				tableOneRowTwo.setHeight(11);
				k = 0;
				for (String values : datas) {
					setCellText(tableOneRowTwo.getCell(k), values, "FFFFFF", 21);
					k++;
				}
				name = excelmap.get(i).get(0);
				fos = new FileOutputStream(output + name + ".doc");
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

	public static void main(String[] args) {
			
		String filePath = "testfiles/templete_all.docx";
		String path1 = "testfiles/ysb_final.xls";
		String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
		String output = "testfiles/";
		String tablePath = "testfiles/tables.xls";
		String excelPath = "testfiles/test.xls";
		try {
			BudgetWriter1 xwpf = new BudgetWriter1( path1,  path2,  filePath,  tablePath,  excelPath,
					 output);
			xwpf.testReadByDoc();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}