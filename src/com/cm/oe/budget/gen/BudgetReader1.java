package com.cm.oe.budget.gen;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
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

public class BudgetReader1 {
	//TODO:添加资源关闭语句
	private String excelPath = "testfiles/ysb_final.xls";
	private String excelPath2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private String path1 = "";
	private String path2 = "";
	
	public BudgetReader1(String path1, String path2){
		this.path1 = path1;
		this.path2 = path2;
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
			if(cell.getCellType()==	HSSFCell.CELL_TYPE_FORMULA){
				cell.setCellFormula(cell.toString());
				value = e.evaluate(cell).getNumberValue();
			}
			if(i==0){
				ids.put(Integer.toString(i), cell.toString());
			}else if(i>0&&cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA){
				ids.put(cell.toString(), Double.toString(value));
			}
		}
		String zh = "";		
		for(int j=3; j<sheet.getPhysicalNumberOfRows(); j++){
			zh="";
			String valuesss = "";
			row = sheet.getRow(j);
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
				if(cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA){
					cell.setCellFormula(cell.toString()); 
					if(cell.toString().contains("IF")){
						valuesss = e.evaluate(cell).getStringValue();
						values.put(Integer.toString(i+1), valuesss);
					}else{
						value = e.evaluate(cell).getNumberValue();
						values.put(Integer.toString(i+1), Double.toString(value));
					}
				}else{
					values.put(Integer.toString(i+1), cell.toString());
				}
			}
			if(zh!=""||zh.length()>0){
				datas.put(zh, values);
			}
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
					values_10 = (int)cell.getNumericCellValue();
					lists.add(Integer.toString(values_10));
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
		String zh = getZhFrom4GYsb();
		System.out.println(zh);
		String key_index = "";
		String values_inner = "";
		Map<String, String> map_data = allDatas.get(zh);
		for(String keys:b3Data.keySet()){
			values_inner = "";
			List<String> values = new ArrayList<String>();
			key_index = b3Data.get(keys).get(1);
			
			values_inner = map_data.get(key_index);
			if(values_inner.length()!=0&&!values_inner.equals("0.0")){
				values.add(b3Data.get(keys).get(0));
				values.add(values_inner);
				results.put(keys, values);
			}
		}
		return results;
	}
	
	
	public String getZhFrom4GYsb() throws IOException{
		//预算表中  第三行 B列的名称为： 单项工程名称:SXZH001TL新建、共址2G、共址其他运营商的(F)（D)宏站基站
		//path1: 预算表   path2 3g4g基础信息
		String results = "";
		String result2 = "";
		FileInputStream fise = new FileInputStream(path1);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(path2));
		
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
			results = e.evaluate(cell2).getStringValue();
		}
		//TL TLD TLFD    SXZH017TL   前面有7位
		if(results.contains("TLFD")){
			int end = results.indexOf("TLFD");
			end = end+4;
			int begin = end-10;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TLD")){
			int end = results.indexOf("TLD");
			end = end+3;
			int begin = end-10;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TL")){
			int end = results.indexOf("TL");
			end = end+2;
			int begin = end-10;
			result2 = results.substring(begin+1, end);
		}
		wb2.close();
		wb.close();
		fise.close();
		System.out.println(result2);
		return result2;
	}
	
	public void readExcel() throws IOException{
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(excelPath2));
		
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);
		HSSFFormulaEvaluator e2= new HSSFFormulaEvaluator(wb2);
		
		String [] strArray = new String[2];
		strArray[0] = "ysb_final.xls";
		strArray[1] = "3G4G工程基站预算基础信息表.xls";
		HSSFFormulaEvaluator[] evals = new HSSFFormulaEvaluator[2];
		evals[0] = e;
		evals[1] = e2;
		HSSFFormulaEvaluator.setupEnvironment(strArray, evals); 
		Sheet sheet3 = wb.getSheetAt(7);
		Row r = null;
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
					values = e.evaluate(cell).getStringValue();
				}		
			}
			if(flag){
				break;
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
		String path1 = "testfiles/ysb_final.xls";
		String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
		BudgetReader1 ub = new BudgetReader1(path1, path2);
		try {
			Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
			Map<String, List<String>> data_b3 = ub.getB3(path1);
			Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path1, path2);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
