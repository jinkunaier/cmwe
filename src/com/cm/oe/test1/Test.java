package com.cm.oe.test1;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class Test {  
      
    public static void main(String[] args) throws Exception {  
          
        Map<String, Object> param = new HashMap<String, Object>();  
        param.put("${name}", "huangqiqing");  
        param.put("${zhuanye}", "信息管理与信息系统");  
        param.put("${school_name}", "山东财经大学");  
                          
        CustomXWPFDocument doc = WordUtil.generateWord(param, "testfiles/testBBU.docx");  
        FileOutputStream fopts = new FileOutputStream("testfiles/write.docx");  
        doc.write(fopts);  
        fopts.close();  
    }  
}  