package com.cm.oe.test;

import java.io.FileInputStream;
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

	public static void main(String[] args) {
		ReplaceMent rm = new ReplaceMent();
		String wordPath = "testfiles/template.doc";
		String excelPath = "testfiles/test.xls";
		String outPath = "testfiles/";
		try {
			rm.replace(excelPath, outPath, wordPath);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private void replace(String excelPath, String outPath, String wordPath) throws Exception {
		Map<Integer, List<String>> excelmap = new HashMap<Integer,List<String>>();
		/****
		 * 读取excel，获得 sheet
		 */
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		Sheet sheet = wb.getSheetAt(0);

		/**
		 * 获得excel中的行数
		 */
		int rowNums = re.rowNumber(wb);
		FileOutputStream fos =null;
		Row r = null;
		String name = null;
		
		/**
		 * 向map中添加每行的数据，并且key为行号，value为每行的数据
		 */
		for (int i = 0; i < rowNums; i++) {
			r = sheet.getRow(i);
			excelmap.put(i, re.getExcelvalues(r));
			
			//fos.add(new FileOutputStream(outPath + name + ".doc"));
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
			List<String> wordList = rw.getWordvalue(range);
			for (int j = 0; j < excelmap.get(i).size(); j++) {
				//range.replaceText(wordList.get(j), excelmap.get(1).get(j));

				range.replaceText(wordList.get(j), excelmap.get(i).get(j));
				//System.out.print(wordList.get(j));
				//System.out.println(excelmap.get(i).get(j));
			}
			name = excelmap.get(i).get(0);
			fos = new FileOutputStream(outPath + name + ".doc");
			doc.write(fos);
			fisw.close();
		}
		
		/**
		 * 关闭所有输入输出流
		 */
		fos.close();
		wb.close();
		fise.close();
		
	}

	public void close(HSSFWorkbook wb) {
		if (wb != null) {
			try {
				wb.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void close(FileInputStream fis) {
		if (fis != null) {
			try {
				fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void close(FileOutputStream fos) {
		if (fos != null) {
			try {
				fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
