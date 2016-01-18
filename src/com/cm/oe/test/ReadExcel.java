package com.cm.oe.test;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReadExcel {

	public int rowNumber(HSSFWorkbook wb) {
		Sheet sht = wb.getSheetAt(0);
		return sht.getPhysicalNumberOfRows();
	}

	public List<String> getExcelvalues(Row row) {
		List<String> line = new ArrayList<String>();
		String cell = null;
		for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
			cell = row.getCell(i).toString();
			line.add(cell);
		}
		return line;
	}
	



}
