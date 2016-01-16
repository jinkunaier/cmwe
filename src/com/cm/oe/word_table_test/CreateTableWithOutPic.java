package com.cm.oe.word_table_test;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
 
public class CreateTableWithOutPic {
	public static void main(String[] args) {
		String outputFile = "D:\\test.doc";
		CustomXWPFDocument document = new CustomXWPFDocument();
		XWPFTable tableOne = document.createTable();
		XWPFTableRow tableOneRowOne = tableOne.getRow(0);
		tableOneRowOne.getCell(0).setText("第1行第1列");
		tableOneRowOne.addNewTableCell().setText("第1行第2列");
		tableOneRowOne.addNewTableCell().setText("第1行第3列");
		tableOneRowOne.addNewTableCell().setText("第1行第4列");
		tableOneRowOne.addNewTableCell().setText("第1行第5列");
		tableOneRowOne.addNewTableCell().setText("第1行第6列");
		tableOneRowOne.addNewTableCell().setText("第1行第7列");
		tableOneRowOne.addNewTableCell().setText("第1行第8列");
		tableOneRowOne.addNewTableCell().setText("第1行第10列");
		tableOneRowOne.addNewTableCell().setText("第1行第11列");
		tableOneRowOne.addNewTableCell().setText("第1行第12列");
		tableOneRowOne.addNewTableCell().setText("第1行第13列");
		XWPFTableRow tableOneRowTwo = tableOne.createRow();
		tableOneRowTwo.getCell(0).setText("第2行第1列");
		tableOneRowTwo.getCell(1).setText("第2行第2列");
		tableOneRowTwo.getCell(2).setText("第2行第3列");
		tableOneRowTwo.getCell(3).setText("第2行第4列");
		tableOneRowTwo.getCell(4).setText("第2行第5列");
		tableOneRowTwo.getCell(5).setText("第2行第6列");
		tableOneRowTwo.getCell(6).setText("第2行第7列");
		tableOneRowTwo.getCell(7).setText("第2行第8列");
		tableOneRowTwo.getCell(8).setText("第2行第9列");
		tableOneRowTwo.getCell(9).setText("第2行第10列");
		tableOneRowTwo.getCell(10).setText("第2行第11列");
		tableOneRowTwo.getCell(11).setText("第2行第12列");
		XWPFTableRow tableOneRowThree = tableOne.createRow();
		tableOneRowThree.getCell(0).setText("第3行第1列");
		tableOneRowThree.getCell(1).setText("第3行第2列");
		tableOneRowThree.getCell(2).setText("第3行第3列");
		tableOneRowThree.getCell(3).setText("第3行第4列");
		tableOneRowThree.getCell(4).setText("第3行第5列");
		tableOneRowThree.getCell(5).setText("第3行第6列");
		tableOneRowThree.getCell(6).setText("第3行第7列");
		tableOneRowThree.getCell(7).setText("第3行第8列");
		tableOneRowThree.getCell(8).setText("第3行第9列");
		tableOneRowThree.getCell(9).setText("第3行第10列");
		tableOneRowThree.getCell(10).setText("第3行第11列");
		tableOneRowThree.getCell(11).setText("第3行第12列");
		XWPFTableRow tableOneRowFour = tableOne.createRow();
		tableOneRowFour.getCell(0).setText("第4行第1列");
		tableOneRowFour.getCell(1).setText("第4行第2列");
		tableOneRowFour.getCell(2).setText("第4行第3列");
		tableOneRowFour.getCell(3).setText("第4行第4列");
		tableOneRowFour.getCell(4).setText("第4行第5列");
		tableOneRowFour.getCell(5).setText("第4行第6列");
		tableOneRowFour.getCell(6).setText("第4行第7列");
		tableOneRowFour.getCell(7).setText("第4行第8列");
		tableOneRowFour.getCell(8).setText("第4行第9列");
		tableOneRowFour.getCell(9).setText("第4行第10列");
		tableOneRowFour.getCell(10).setText("第4行第11列");
		tableOneRowFour.getCell(11).setText("第4行第12列");
		FileOutputStream fOut;
		try {
			fOut = new FileOutputStream(outputFile);
			document.write(fOut); 
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}
}
