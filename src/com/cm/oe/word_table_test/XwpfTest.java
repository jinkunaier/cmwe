package com.cm.oe.word_table_test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;

public class XwpfTest {
	  private final static String filePath = "testfiles/templete_all.docx";
	   /**
	    * 通过XWPFDocument对内容进行访问。对于XWPF文档而言，用这种方式进行读操作更佳。
	    * @throws Exception
	    */
	   @Test
	   public void testReadByDoc() throws Exception {
	      InputStream is = new FileInputStream(filePath);
	      OutputStream os = new FileOutputStream("testfiles/templete_all2.docx");  
	      XWPFDocument doc = new XWPFDocument(is);
//	      List<XWPFParagraph> paras = doc.getParagraphs();
//	      for (XWPFParagraph para : paras) {
//	         //当前段落的属性
////	       CTPPr pr = para.getCTP().getPPr();
//	         System.out.println(para.getText());
//	      }
	      //获取文档中所有的表格
	      List<XWPFTable> tables = doc.getTables();
	      List<XWPFTableRow> rows;
	      List<XWPFTableCell> cells;
	      for (XWPFTable table : tables) {
	         //表格属性
//	       CTTblPr pr = table.getCTTbl().getTblPr();
	         //获取表格对应的行
	         rows = table.getRows();
	         for (XWPFTableRow row : rows) {
	            //获取行对应的单元格
	            cells = row.getTableCells();
	            for (XWPFTableCell cell : cells) {
	                System.out.println(cell.getText());;
	            }
	         }
	 		XWPFTableRow tableOneRowTwo = table.createRow();
	 		tableOneRowTwo.getCell(0).setText("第2行第1列");
	      }
	      doc.write(os);
	      os.flush();
	      this.close(is);
	      os.close();
	   }
	  
	   /**
	    * 关闭输入流
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
	   
	   public static void main(String[] args){
		   try {
			XwpfTest xwpf = new XwpfTest();
			xwpf.testReadByDoc();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	   }
	  
	}