package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.jxcell.View;

public class UpdateBudget {
	//TODO:添加资源关闭语句
	private ReadExcel re;
	private String excelPath = "testfiles/ysb_final.xls";
	private String excelPath2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private String excelPath3 = "testfiles/无线主设备设计表一表五.xls";
	private String excelPath4 = "testfiles/testexcel.xls";
	View m_view; 
	
	public UpdateBudget(){
		re = new ReadExcel();
	}
	
	public Map<String, Map<String, String>> get3G4Gjcxx(String path) throws FileNotFoundException, IOException{
		//获取3G4G工程预算基础信息表里面基础信息sheet中的  工程项目 以及对应的工程信息  外层map的键对应的是站号， 内层map对应的信息是id及取值
		//在基础信息表中站名不能重复
		//不能在两行之间存在空行
		HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(path));
		Sheet sheet = wb.getSheetAt(1);
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);
		Map<String,String> ids = new LinkedHashMap<String,String>();
		Map<String, Map<String, String>> datas = new LinkedHashMap<String, Map<String, String>>();
		//获取第二行的键和值的对应关系
		Row row = sheet.getRow(2);
		Cell cell = null;		
		for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
			cell = row.getCell(i);
			double value = 0;
			String values = "";
			if(cell.getCellType()==	HSSFCell.CELL_TYPE_FORMULA){
				cell.setCellFormula(cell.toString());
//				if(cell.toString().contains("IF")){
//					values = e.evaluate(cell).getStringValue();
//				}else{
//					value = e.evaluate(cell).getNumberValue();
//				}
				value = e.evaluate(cell).getNumberValue();
			}
			if(i==0){
				ids.put(Integer.toString(i), cell.toString());
			}else if(i>0&&cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA){
				ids.put(cell.toString(), Double.toString(value));
//				if(cell.toString().contains("IF")){
//					ids.put(Integer.toString(i), values);
//				}else{
//					ids.put(Integer.toString(i), Double.toString(value));	
//				}
			}
		}
//		for(String key:ids.keySet()){
//			System.out.println("key="+key+", value is="+ids.get(key));
//		}
		String zh = "";		
		for(int j=3; j<sheet.getPhysicalNumberOfRows(); j++){
			zh="";
			String valuesss = "";
			row = sheet.getRow(j);
			Cell cell2 = null;	
			Map<String, String> values = new LinkedHashMap<String, String>();
			if(row.getCell(3)==null){
				continue;
			}
			zh = row.getCell(3).toString();
			if(zh==""||zh.length()==0){
				continue;
			}
			
			for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
				cell = row.getCell(i);
				double value = 0.0;
				if(i==3){
					zh = cell.toString();
				}
//				System.out.println("j-=="+j+", i=="+i+", zh=="+zh);
				
				if(cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA){
					cell.setCellFormula(cell.toString()); 
					if(cell.toString().contains("IF")){
						valuesss = e.evaluate(cell).getStringValue();
						values.put(Integer.toString(i+1), valuesss);
					}else{
						value = e.evaluate(cell).getNumberValue();
						values.put(Integer.toString(i+1), Double.toString(value));
					}
//					value = e.evaluate(cell).getNumberValue();

				}else{
					values.put(Integer.toString(i+1), cell.toString());
				}
			}
			if(zh!=""||zh.length()>0){
				datas.put(zh, values);
			}
//			for(String key : datas.keySet()){
//				System.out.println("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"+key);
//				for(String key2:datas.get(key).keySet()){
//					System.out.println("key is==="+key2+", key2 value is==="+datas.get(key).get(key2));
//				}
//			}
		}
		wb.close();
		return datas;
	}
	
	
	public Map<String, List<String>> getB3(String path) throws FileNotFoundException, IOException{
		//4G工程预算输出表中的格式不能改变，表中的项目名称不得有重复，否则程序出错
		//获取B3甲表中的 项目名称  单位  以及序号
		HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(path));
		Map<String, List<String>> datas = new LinkedHashMap<String, List<String>>();
		Sheet sheet = wb.getSheetAt(7);
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);
		Row r = null;
		String key = "";
		int values_10 = 0;
		for(int i=7; i<sheet.getPhysicalNumberOfRows(); i++){
			r = sheet.getRow(i);
			String scell = "";
			Cell cell = null;
			key = "";
			boolean flag = false;
			List<String> lists = new ArrayList<String>();
			for (int j = 3; j <=10; j++) {
				if(j==5||j==6||j==7||j==8||j==9){
					continue;
				}
				cell = r.getCell(j);
				scell = cell.toString();
				if(cell==null||scell==""){
					flag = true;
					break;
				}
				if(j==3){
					key = scell;
				}else if(j==10){
//					System.out.println(scell);
					values_10 = (int)cell.getNumericCellValue();
					lists.add(Integer.toString(values_10));
//					System.out.println(Integer.toString(values_10));
				}else{
					lists.add(scell);
				}
			}
			if(!flag){
				datas.put(key, lists);
			}
			if(flag){
				break;
			}
		}
		wb.close();
		return datas;
	}
	
	public Map<String, List<String>> getB3JData(Map<String, Map<String, String>> allDatas, Map<String, List<String>> b3Data, String path1, String path2) throws IOException{
		//通过遍历从B3表中读取的信息，结合从3G4G表中读取的数据，生成最终的真实数据
		Map<String, List<String>> results = new LinkedHashMap<String,  List<String>>();
		String zh = getZhFrom4GYsb(path1, path2);
		System.out.println(zh);
		String key_index = "";
		String values_inner = "";
		Map<String, String> map_data = allDatas.get(zh);
//		for(String key:map_data.keySet()){
//			System.out.println("key="+key+", value is="+map_data.get(key));
//			System.out.println(map_data.get(key).length());
//		}
		double key_value = 0.0;
		for(String keys:b3Data.keySet()){
			values_inner = "";
			List<String> values = new ArrayList<String>();
			key_index = b3Data.get(keys).get(1);
			
			values_inner = map_data.get(key_index);
//			System.out.println("key_index ==="+key_index+", values ==="+values_inner);
			if(values_inner.length()!=0&&!values_inner.equals("0.0")){
//				System.out.println("++++++++++++++++++++++++"+"key_index ==="+key_index+", values ==="+values_inner);
				values.add(b3Data.get(keys).get(0));
				values.add(values_inner);
				results.put(keys, values);
			}
		}
		
		for(String key : results.keySet()){
			System.out.println(key);
			for(String key2:results.get(key)){
				System.out.print(key2+",  ");
			}
			System.out.println("-----------------");
		}
		return results;
	}
	
	
	public String getZhFrom4GYsb(String path1, String path2) throws IOException{
		//预算表中  第三行 B列的名称为： 单项工程名称:SXZH001TL新建、共址2G、共址其他运营商的(F)（D)宏站基站
		//path1: 预算表   path2 3g4g基础信息
		String results = "";
		String result2 = "";
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(excelPath2));
		
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);
		HSSFFormulaEvaluator e2= new HSSFFormulaEvaluator(wb2);
		
		//TODO: 将此处的文件名替换为从参数读取的文件名
		String [] strArray = new String[2];
		strArray[0] = "ysb_final.xls";
		strArray[1] = "3G4G工程基站预算基础信息表.xls";
		HSSFFormulaEvaluator[] evals = new HSSFFormulaEvaluator[2];
		evals[0] = e;
		evals[1] = e2;
		HSSFFormulaEvaluator.setupEnvironment(strArray, evals); 
		
		Sheet sheet = wb.getSheetAt(7);
		Row r = null;
		r = sheet.getRow(2);
		Cell cell2 = r.getCell(1);
		if(cell2.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
//			System.out.println(cell2.toString());
			results = e.evaluate(cell2).getStringValue();
//			System.out.println(results);
		}
		if(results.contains("新建")){
			int begin = results.indexOf(":");
			int end = results.indexOf("新建");
			result2 = results.substring(begin+1, end);
		}else if(results.contains("共建")){
			int begin = results.indexOf(":");
			int end = results.indexOf("共建");
			result2 = results.substring(begin+1, end);
		}
		wb2.close();
		wb.close();
		fise.close();
		return result2;
	}
	
	public void readExcel() throws IOException{
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(excelPath2));
		HSSFWorkbook wb3 = new HSSFWorkbook(new FileInputStream(excelPath3));
		POIFSFileSystem poiexcel2 = new POIFSFileSystem(new FileInputStream(excelPath2));
		POIFSFileSystem poiexcel3 = new POIFSFileSystem(new FileInputStream(excelPath3));
		
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);
		HSSFFormulaEvaluator e2= new HSSFFormulaEvaluator(wb2);
		HSSFFormulaEvaluator e3= new HSSFFormulaEvaluator(wb3);
		
		String [] strArray = new String[3];
		strArray[0] = "ysb_final.xls";
		strArray[1] = "3G4G工程基站预算基础信息表.xls";
		strArray[2] = "无线主设备设计表一表五.xls";
		HSSFFormulaEvaluator[] evals = new HSSFFormulaEvaluator[3];
		evals[0] = e;
		evals[1] = e2;
		evals[2] = e3;
		HSSFFormulaEvaluator.setupEnvironment(strArray, evals); 
		//col K---BU
		
		
		Sheet sheet3 = wb.getSheetAt(7);
		Row r = null;
//		r = sheet3.getRow(2);
//		Cell cell2 = r.getCell(1);
//		if(cell2.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
//			System.out.println(cell2.toString());
//			String value = e.evaluate(cell2).getStringValue();
//			System.out.println(value);
//		}
//		
//		r = sheet3.getRow(0);
//		re.printExcelvalues(r);
//		cell2 = r.getCell(0);
//		if(cell2.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
//			System.out.println(cell2.toString());
//			String value = e.evaluate(cell2).getStringValue();
//			System.out.println(value);
//		}
		
		for(int i=7; i<sheet3.getPhysicalNumberOfRows(); i++){
			r = sheet3.getRow(i);
			String scell = "";
			Cell cell = null;
			boolean flag = false;
			String values = "";
			for (int j = 3; j <=5; j++) {
				cell = r.getCell(j);
				scell = cell.toString();
				if(cell==null||scell==""){
					flag = true;
					break;
				}
				System.out.println(scell);
				if(cell.getCellType()==	HSSFCell.CELL_TYPE_FORMULA){
//					int value = e.evaluateFormulaCell(cell);
//					cell.setCellFormula(cell.toString()); 
//					int value = e.evaluateFormulaCell(cell);
//					double value = e.evaluate(cell).getNumberValue();
					values = e.evaluate(cell).getStringValue();
					System.out.println(values);
				}		
			//e.clearAllCachedResultValues();	
			}
			if(flag){
				break;
			}
			//re.printExcelvalues(r);
		}
	}
	
	
	public void test_Excel_struct() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(excelPath2));
		Sheet sheet3 = wb.getSheetAt(1);
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);

//		Row row = sheet3.getRow(2);
		Row row = sheet3.getRow(3);
//		re.printExcelvalues(row);
		// row and col begin from index 0
		Cell cell = null;
//		cell = row.getCell(8)
		for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
			cell = row.getCell(i);

			if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
				System.out.println("row num is ====="+(i+1));
				System.out.println(cell.toString());
			}else if(cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
				String value = "";
//				System.out.println(cell.toString());
				cell.setCellFormula (cell.toString()); 
//				int value = e.evaluateFormulaCell(cell);
				 
				if(cell.toString().contains("IF")){
					value = e.evaluate(cell).getStringValue();
					if(value.length()>0){
						System.out.println("row num is ====="+(i+1));
						System.out.println(value);
					}
				}else{
					double value2 = e.evaluate(cell).getNumberValue();
					if(value2>0){
						System.out.println("row num is ====="+(i+1));
						System.out.println(value2);	
					}
				}
				
			}else if(cell.toString().length()>0){
				System.out.println("row num is ====="+(i+1));
				System.out.println(cell.toString());
			}

		}

	}
	
	public void printMapValue(Map<String, List<String>> datas){
		List<String> vals = null;
		for(String key : datas.keySet()) {
			System.out.println("key= "+ key);
			vals = datas.get(key);
			for(String temp:vals){
				System.out.print(temp+", ");
			}
			System.out.println();
		}
	}
	
	public static void main(String[] args) {
		UpdateBudget ub = new UpdateBudget();
		try {
			String path1 = "testfiles/ysb_final.xls";
			String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
//			ub.readExcel();
//			ub.readExcel_jxl();
//			ub.test_Excel_struct();
//			Map<String, List<String>> datas = new LinkedHashMap<String, List<String>>();
//			datas = ub.getB3("testfiles/ysb_final.xls");
//			ub.printMapValue(datas);
//			System.out.println(datas.size());
			
			
//			Map<String, Map<String, String>> datas = new LinkedHashMap<String, Map<String, String>>();
//			datas = ub.get3G4Gjcxx("testfiles/3G4G工程基站预算基础信息表.xls");
////			ub.printMapValue(datas);
//			System.out.println(datas.size());
			
			Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
			Map<String, List<String>> data_b3 = ub.getB3(path1);
			Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path1, path2);
//			ub.test_Excel_struct();
//			ub.readExcel();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
