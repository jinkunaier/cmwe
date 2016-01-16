package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReplaceMent {
	ReadExcel re = new ReadExcel();
	ReadWord rw = new ReadWord();
	ReadExcelTable ret = new ReadExcelTable();
	public static void main(String[] args) {
		ReplaceMent rm = new ReplaceMent();
		String wordPath = "testfiles/template.doc";
		String excelPath = "testfiles/test.xls";
		String tablePath = "testfiles/tables.xls";
		String outPath = "testfiles/";
		try {
			rm.replace(excelPath, outPath, wordPath,tablePath);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	// replace optional file
	private void replacefile(String excelPath, String outpath, String wordPath, String wordContent) throws IOException{
		//创建一个map，其中存储了excel每一行对应的信息
		Map<Integer, List<String>> excelmap = new HashMap<Integer, List<String>>();
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		Sheet sheet = wb.getSheetAt(0);
		/**
		 * 获得excel中的行数
		 */
		int rowNums = re.rowNumber(wb);
		FileOutputStream fos = null;
		Row r = null;
		String name = null;

		/**
		 * 向map中添加每行的数据，并且key为行号，value为每行的数据
		 */
		for (int i = 0; i < rowNums; i++) {
			r = sheet.getRow(i);
			excelmap.put(i, re.getExcelvalues(r));
		}
		
		
	}

	public void replace(String excelPath, String outPath, String wordPath,String tablePath) throws Exception {
		//创建一个map，其中存储了excel每一行对应的信息
		Map<Integer, List<String>> excelmap = new HashMap<Integer, List<String>>();
		//创建一个tablemap 读取厂家对应的列信息
		Map<Integer, List<String>> tablemap = ret.readTableinExcel(tablePath,excelPath);
		/****
		 * 读取excel，获得 sheet
		 *  excelPath : excel路径
		 *  outpath : 文件输出路径
		 *  wordPath : word路径
		 *   
		 */
		
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		Sheet sheet = wb.getSheetAt(0);

		/**
		 * 获得excel中的行数
		 */
		int rowNums = re.rowNumber(wb);
		FileOutputStream fos = null;
		Row r = null;
		String name = null;

		/**
		 * 向map中添加每行的数据，并且key为行号，value为每行的数据
		 */
		for (int i = 0; i < rowNums; i++) {
			r = sheet.getRow(i);
			excelmap.put(i, re.getExcelvalues(r));
		}
		/**
		 * 以excel中每行的数据生成一个新的doc文件。
		 */
		for (int i = 0; i < rowNums; i++) {
			/**
			 * 读取word，获得所有标记的文字
			 */
			FileInputStream fisw = new FileInputStream(wordPath);
			HWPFDocument doc = new HWPFDocument(fisw);
			Range range = rw.getRange(doc);
			/**
			 * 读取word，获得段落里${}标记的文字
			 */
			List<String> wordParalist = rw.getParavalue(range);
			/**
			 * 读取word，获得表格里#{}标记的文字
			 */
			List<String> wordTablelist= rw.getTablevalue(range);
			/**
			 * 替换段落里的文字
			 */
			for (int j = 0; j < excelmap.get(i).size(); j++) {
				range.replaceText(wordParalist.get(j), excelmap.get(i).get(j));
			}
			/**
			 * 替换表格里的文字
			 */
			for (int j=0;j < tablemap.get(0).size();j++){
				range.replaceText(wordTablelist.get(j), tablemap.get(i).get(j));
			}
			name = excelmap.get(i).get(0);
			fos = new FileOutputStream(outPath + name + ".doc");
			doc.write(fos);
			wb.close();
			fisw.close();
		}

		/**
		 * 关闭所有输入输出流
		 */
		fos.close();
		wb.close();
		fise.close();

	}
	
}
