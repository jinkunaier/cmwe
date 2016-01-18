package com.cm.oe.test1;


import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class testttt {
	public void word(String file, String newFile) {
		try {
			OPCPackage pack = POIXMLDocument.openPackage(file);
			XWPFDocument doc = new XWPFDocument(pack);
			List<XWPFParagraph> paragraphs = doc.getParagraphs();
			System.out.println(paragraphs.size());
			for (XWPFParagraph tmp : paragraphs) {
				System.out.println(tmp.getParagraphText());
				List<XWPFRun> runs = tmp.getRuns();
				for (XWPFRun aa : runs) {
					System.out.println("XWPFRun-Text:" + aa.getText(0));
					if ("$name".equals(aa.getText(0))) {
						aa.setText("必先利", 0);
					}else if("$$$$$".equals(aa.getText(0))){
						aa.setText("xupt",0);
					}else if("$school".equals(aa.getText(0))){
						aa.setText("西安",0);
					}
				}
			}

			FileOutputStream fos = new FileOutputStream(newFile);
			doc.write(fos);
			fos.flush();
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		testttt tools = new testttt();
		try {
			tools.word("testfiles/testBBU.docx", "testfiles/write.docx");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
