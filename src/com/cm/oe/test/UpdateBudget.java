package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
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
	private ReadExcel re;
	private String excelPath = "testfiles/ysb_final.xls";
	private String excelPath2 = "testfiles/3G4G工程基站预算基础信息表.xls";
	private String excelPath3 = "testfiles/无线主设备设计表一表五.xls";
	View m_view; 
	
	public UpdateBudget(){
		re = new ReadExcel();
	}
	
	public Map<String, List<String>> getB3(Sheet sheet){
		//4G工程预算输出表中的格式不能改变，表中的项目名称不得有重复，否则程序出错
		Map<String, List<String>> datas = new HashMap<String, List<String>>();
		Row r = null;
		String key = "";
		for(int i=7; i<sheet.getPhysicalNumberOfRows(); i++){
			r = sheet.getRow(i);
			String scell = "";
			Cell cell = null;
			key = "";
			boolean flag = false;
			List<String> lists = new ArrayList<String>();
			for (int j = 3; j <=4; j++) {
				cell = r.getCell(j);
				scell = cell.toString();
				if(cell==null||scell==""){
					flag = true;
					break;
				}
				if(j==3){
					key = scell;
				}else{
					lists.add(scell);
				}
			}
			datas.put(key, lists);
			if(flag){
				break;
			}
		}
		return datas;
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
		//K---BU
//		// Create a FormulaEvaluator to use
//		FormulaEvaluator mainWorkbookEvaluator = wb.getCreationHelper().createFormulaEvaluator();
//
//		// Track the workbook references
//		Map<String,FormulaEvaluator> workbooks = new HashMap<String, FormulaEvaluator>();
//		// Add this workbook
//		workbooks.put("ysb_final.xls", mainWorkbookEvaluator);
//		// Add two others
//		workbooks.put("3G4G工程基站预算基础信息表.xls", WorkbookFactory.create(poiexcel2).getCreationHelper().createFormulaEvaluator());
//		workbooks.put("无线主设备设计表一表五.xls", WorkbookFactory.create(poiexcel3).getCreationHelper().createFormulaEvaluator());
//
//		// Attach them
//		mainWorkbookEvaluator.setupReferencedWorkbooks(workbooks);
//
//		// Evaluate
//		mainWorkbookEvaluator.evaluateAll();
		
		
		Sheet sheet3 = wb.getSheetAt(7);
		Row r = null;
		r = sheet3.getRow(2);
		Cell cell2 = r.getCell(1);
		if(cell2.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
			System.out.println(cell2.toString());
			String value = e.evaluate(cell2).getStringValue();
			System.out.println(value);
		}
		for(int i=7; i<sheet3.getPhysicalNumberOfRows(); i++){
			r = sheet3.getRow(i);
			String scell = "";
			Cell cell = null;
			boolean flag = false;
			for (int j = 3; j <=5; j++) {
				cell = r.getCell(j);
				scell = cell.toString();
				if(cell==null||scell==""){
					flag = true;
					break;
				}
				System.out.println(scell);
				if(cell.getCellType()==	HSSFCell.CELL_TYPE_FORMULA){
					int value = e.evaluateFormulaCell(cell);
					System.out.println(value);
				}		
			//e.clearAllCachedResultValues();	
			}
			if(flag){
				break;
			}
			//re.printExcelvalues(r);
		}
	}
	
	public static void main(String[] args) {
		UpdateBudget ub = new UpdateBudget();
		try {
			ub.readExcel();
//			ub.readExcel_jxl();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
