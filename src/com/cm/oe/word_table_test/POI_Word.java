package com.cm.oe.word_table_test;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
  
public class POI_Word{  
	private final static String filePath = "testfiles/templete_all.doc";
    public  static void main(String[] args){  
        try {  
            String[] s=new String[100];  
            List<String> lists = new ArrayList<String>();
            FileInputStream in=new FileInputStream(filePath);  
            POIFSFileSystem pfs=new POIFSFileSystem(in);  
            HWPFDocument hwpf=new HWPFDocument(pfs);  
            Range range =hwpf.getRange();  
            TableIterator it=new TableIterator(range);  
//            int index=0;  
            while(it.hasNext()){  
                Table tb=(Table)it.next();  
                for(int i=0;i<tb.numRows();i++){  
//                    System.out.println("Numrows :"+tb.numRows());  
                    TableRow tr=tb.getRow(i);  
                    for(int j=0;j<tr.numCells();j++){  
                        System.out.println("numCells :"+tr.numCells());  
                        System.out.println("j   :"+j);  
                        TableCell td=tr.getCell(j);  
                        for(int k=0;k<td.numParagraphs();k++){  
                            System.out.println("numParagraphs :"+td.numParagraphs());  
                            Paragraph para=td.getParagraph(k);  
//                            s[index]=para.text().trim();  
                            lists.add(para.text().trim());
//                            index++;  
                        }  
                    }  
                } 
//        		XWPFTableRow tableOneRowTwo = tb.createRow();
            }  
//          System.out.println(s.toString());  
            for(int i=0;i<lists.size();i++){  
                System.out.println(lists.get(i));  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
}









//FileInputStream fileInputStream = new FileInputStream( soureFile);
//POIFSFileSystem pfs = new POIFSFileSystem( fileInputStream );
//HWPFDocument hwpf = new HWPFDocument(pfs);// make a HWPFDocument object
//
//OutputStream output = new FileOutputStream( targetFile );
//hwpf.write(output);// write to the target file
//output.close();
//
//（2）再word中插入表格。HWPF的情况：
//Table tcDataTable = range.insertTableBefore( (short)column , row);//column and row列数和行数
//tcDataTable.getRow(i).getCell(j).getParagraph(0).getCharacterRun(0).insertBefore("插入i行j列的内容" );
//
//XWPF的情况：
//String outputFile = "D:\\test.doc";
//
//XWPFDocument document = new XWPFDocument();
//
//XWPFTable tableOne = document.createTable();
//
//
//
//
//XWPFTableRow tableOneRowOne = tableOne.getRow(0);
//tableOneRowOne.getCell(0).setText("11");
//XWPFTableCell cell12 =   tableOneRowOne.createCell();
//cell12.setText("12");
////	tableOneRowOne.addNewTableCell().setText("第1行第2列");
////	tableOneRowOne.addNewTableCell().setText("第1行第3列");
////	tableOneRowOne.addNewTableCell().setText("第1行第4列");
//
//XWPFTableRow tableOneRowTwo = tableOne.createRow();
//tableOneRowTwo.getCell(0).setText("21");
//tableOneRowTwo.getCell(1).setText("22");
////	tableOneRowTwo.getCell(2).setText("第2行第3列");
//
//XWPFTableRow tableOneRow3 = tableOne.createRow();
//tableOneRow3.addNewTableCell().setText("31");
//tableOneRow3.addNewTableCell().setText("32");
//
//FileOutputStream fOut;
//
//try {
//fOut = new FileOutputStream(outputFile);
//
//document.write(fOut); 
//fOut.flush();
//// 操作结束，关闭文件
//fOut.close();
//} catch (Exception e) {
//e.printStackTrace();
//} 
