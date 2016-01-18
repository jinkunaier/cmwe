package com.cm.oe.test1;  
  
import java.io.ByteArrayInputStream;  
import java.io.IOException;  
import java.io.InputStream;  
import java.util.Iterator;  
import java.util.List;  
import java.util.Map;  
import java.util.Map.Entry;  
import org.apache.poi.POIXMLDocument;  
import org.apache.poi.openxml4j.opc.OPCPackage;  
import org.apache.poi.xwpf.usermodel.XWPFParagraph;  
import org.apache.poi.xwpf.usermodel.XWPFRun;  
import org.apache.poi.xwpf.usermodel.XWPFTable;  
import org.apache.poi.xwpf.usermodel.XWPFTableCell;  
import org.apache.poi.xwpf.usermodel.XWPFTableRow;  
  
/** 
 * 适用于word 2007 
 * poi 版本 3.7 
 */  
public class WordUtil {  
  
    /** 
     * 根据指定的参数值、模板，生成 word 文档 
     * @param param 需要替换的变量 
     * @param template 模板 
     */  
    public static CustomXWPFDocument generateWord(Map<String, Object> param, String template) {  
        CustomXWPFDocument doc = null;  
        try {  
            OPCPackage pack = POIXMLDocument.openPackage(template);  
            doc = new CustomXWPFDocument(pack);  
            if (param != null && param.size() > 0) {  
                  
                //处理段落  
                List<XWPFParagraph> paragraphList = doc.getParagraphs();  
                processParagraphs(paragraphList, param, doc);  
                  
                //处理表格  
                Iterator<XWPFTable> it = doc.getTablesIterator();  
                while (it.hasNext()) {  
                    XWPFTable table = it.next();  
                    List<XWPFTableRow> rows = table.getRows();  
                    for (XWPFTableRow row : rows) {  
                        List<XWPFTableCell> cells = row.getTableCells();  
                        for (XWPFTableCell cell : cells) {  
                            List<XWPFParagraph> paragraphListTable =  cell.getParagraphs();  
                            processParagraphs(paragraphListTable, param, doc);  
                        }  
                    }  
                }  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return doc;  
    }  
    /** 
     * 处理段落 
     * @param paragraphList 
     */  
    public static void processParagraphs(List<XWPFParagraph> paragraphList,Map<String, Object> param,CustomXWPFDocument doc){  
        if(paragraphList != null && paragraphList.size() > 0){  
            for(XWPFParagraph paragraph:paragraphList){  
                List<XWPFRun> runs = paragraph.getRuns();  
                for (XWPFRun run : runs) {  
                    String text = run.getText(0);  
                    System.out.println(text);
                    if(text != null){  
                        boolean isSetText = false;  
                        for (Entry<String, Object> entry : param.entrySet()) {  
                            String key = entry.getKey();  
                            System.out.println(text.indexOf(key));
                            if(text.indexOf(key) == -1){  
                                isSetText = true;  
                                Object value = entry.getValue();  
                                if (value instanceof String) {//文本替换  
                                    text = text.replace(key, value.toString());  
                                } 
                            }  
                        }  
                        if(isSetText){  
                            run.setText(text,0);  
                        }  
                    }  
                }  
            }  
        }  
    }  
}  