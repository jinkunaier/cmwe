package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReadExcelTable {
	ReadWord rw = new ReadWord();
	ReadExcel re = new ReadExcel();
	
	public Map<Integer, List<String>> readTableinExcel(String tablePath,String excelPath) throws IOException {
		FileInputStream fist = new FileInputStream(tablePath);
		HSSFWorkbook wbt = new HSSFWorkbook(fist);
		Sheet sht_t = wbt.getSheetAt(0);
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wbe = new HSSFWorkbook(fise);
		Sheet sht_e = wbe.getSheetAt(0);
		List<Row> rowe = new ArrayList<Row>();
		for (int i = 0; i < re.rowNumber(wbe); i++) {
			rowe.add(sht_e.getRow(i));
		}
		Row rowt = null;
		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();

		for (int i = 0; i < re.rowNumber(wbe); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (rowe.get(i).getCell(4).toString().equals("华为")) {
				rowt = sht_t.getRow(0);
				for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
					tablevalues.add(rowt.getCell(j).toString());
				}
				tableMap.put(i, tablevalues);
			} else if (rowe.get(i).getCell(4).toString().equals("大唐")) {
				rowt = sht_t.getRow(3);
				for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
					tablevalues.add(rowt.getCell(j).toString());
				}
				tableMap.put(i, tablevalues);
			} else if (rowe.get(i).getCell(4).toString().equals("中兴")) {
				rowt = sht_t.getRow(1);
				for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
					tablevalues.add(rowt.getCell(j).toString());
				}
				tableMap.put(i, tablevalues);
			}
		}
		wbe.close();
		fise.close();
		wbt.close();
		fist.close();
		return tableMap;

	}

}
