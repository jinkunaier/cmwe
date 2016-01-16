package com.cm.oe.test;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReadExcel {

	public int rowNumber(HSSFWorkbook wb) {
		//获取第0个sheet页
		Sheet sht = wb.getSheetAt(0);
		//返回该sheet页面中的记录行数
		return sht.getPhysicalNumberOfRows();
	}

	public List<String> getExcelvalues(Row row) {
		List<String> line = new ArrayList<String>();
		String cell = null;
		for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
			//从第0列开始遍历，将结果存入到一个list当中
			cell = row.getCell(i).toString();
			line.add(cell);
		}
		return line;
	}
	
	public void printExcelvalues(Row row) {
		String cell = null;
		for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
			cell = row.getCell(i).toString();
			System.out.println(cell);
		}
	}
	



}
