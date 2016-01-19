package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.cm.oe.ui.MainPathCreate;

public class ReadExcelTable {
	ReadWord rw = new ReadWord();
	ReadExcel re = new ReadExcel();
	MainPathCreate mc = new MainPathCreate();

	public Map<Integer, List<String>> readBBUinExcel(String tablePath, String excelPath) throws IOException {
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

		if (rowe.get(rowe.size() - 1).getCell(0) == null) {
			JOptionPane.showMessageDialog(null, "请尝试删除" + mc.aText.getText() + "文件中数据行以后的多行空白行(或多个空白列)！");
			System.exit(0);
		}
		// System.out.println(rowe.get(rowe.size()-1).getCell(0)==null);
		// System.out.println(rowe.size());
		Row rowt = null;
		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();
		for (int i = 1; i < rowe.size(); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (rowe.get(i).getCell(15).toString().equals("华为")) {
				if (rowe.get(i).getCell(16).toString().equals("DBBP530(BBU3900)")) {
					rowt = sht_t.getRow(0);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的华为BBU设备型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(15).toString().equals("大唐")) {
				if(rowe.get(i).getCell(16).toString().equals("EMB5116")){
					rowt = sht_t.getRow(3);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				}else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的大唐BBU设备型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(15).toString().equals("中兴")) {
				if(rowe.get(i).getCell(16).toString().equals("ZXSDR B8300")){
					rowt = sht_t.getRow(1);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					
					tableMap.put(i, tablevalues);
				}else{
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的中兴BBU设备型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(15).toString().equals("上海贝尔")) {
				if(rowe.get(i).getCell(16).toString().equals("9926 BBU")){
					rowt = sht_t.getRow(2);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				}else{
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的上海贝尔BBU设备型号！");
					System.exit(0);
				}
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的BBU品牌");
				System.exit(0);
			}
		}
		wbe.close();
		fise.close();
		wbt.close();
		fist.close();
		return tableMap;

	}

	public Map<Integer, List<String>> readRRUinExcel(String tablePath, String excelPath) throws IOException {
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
		System.out.println(rowe.size());
		System.out.println(rowe.get(1).getCell(4));
		System.out.println(rowe.get(1).getCell(7));
		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();
		Row rowt = null;
		for (int i = 1; i < rowe.size(); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (rowe.get(i).getCell(17).toString().equals("华为")) {
				if (rowe.get(i).getCell(18).toString().equals("AAU3213")) {
					rowt = sht_t.getRow(4);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("RRU3277")) {
					rowt = sht_t.getRow(5);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("DRRU3168e-fa")) {
					rowt = sht_t.getRow(6);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("RRU3278M")) {
					rowt = sht_t.getRow(7);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("DRRU3172-fad")) {
					rowt = sht_t.getRow(8);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("DRRU3161-fae")) {
					rowt = sht_t.getRow(9);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("BTS3205E")) {
					rowt = sht_t.getRow(10);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("BBOK RRU")) {
					rowt = sht_t.getRow(11);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("Easymacro")) {
					rowt = sht_t.getRow(12);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的华为RRU设备型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(17).toString().equals("中兴")) {
				if (rowe.get(i).getCell(18).toString().equals("ZXSDR R8978 S2600W")) {
					rowt = sht_t.getRow(13);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("ZXSDR R8972E S2600W")) {
					rowt = sht_t.getRow(14);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("ZXSDR R8978 M1920A")) {
					rowt = sht_t.getRow(15);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("ZXSDR R8972E M1920A")) {
					rowt = sht_t.getRow(16);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("ZXSDR R8972E S2300W")) {
					rowt = sht_t.getRow(17);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("ZXSDR R8972E M192023A")) {
					rowt = sht_t.getRow(18);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的中兴RRU设备型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(17).toString().equals("大唐")) {
				if (rowe.get(i).getCell(18).toString().equals("TDRU348FA")) {
					rowt = sht_t.getRow(19);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TDRU348D")) {
					rowt = sht_t.getRow(20);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TDRU342D")) {
					rowt = sht_t.getRow(21);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TDRU342FA")) {
					rowt = sht_t.getRow(22);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TDRU342E")) {
					rowt = sht_t.getRow(23);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TDRU341FAE")) {
					rowt = sht_t.getRow(24);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("mTDRU342D")) {
					rowt = sht_t.getRow(25);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的大唐RRU设备型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(17).toString().equals("上海贝尔")) {
				if (rowe.get(i).getCell(18).toString().equals("TD-RRH8X20-25A")) {
					rowt = sht_t.getRow(26);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TD-RRH2×40-25A")) {
					rowt = sht_t.getRow(27);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TD-RRH8x10-1935")) {
					rowt = sht_t.getRow(28);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TD-RRH2x60-1935")) {
					rowt = sht_t.getRow(29);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("TD-RRH2X50-2350")) {
					rowt = sht_t.getRow(30);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(18).toString().equals("9768 MRO B38 TD-LTE 2x5W")) {
					rowt = sht_t.getRow(31);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的上海贝尔RRU设备的型号！");
					System.exit(0);
				}
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的RRU品牌！");
				System.exit(0);
			}
		}
		wbe.close();
		fise.close();
		wbt.close();
		fist.close();
		return tableMap;
	}

	public Map<Integer, List<String>> readAntennaIntables(String tablePath, String excelPath) throws IOException {
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

		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();
		Row rowt = null;
		for (int i = 1; i < rowe.size(); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (rowe.get(i).getCell(19).toString().equals("华为")) {
				if (rowe.get(i).getCell(20).toString().equals("ATD-")) {
					rowt = sht_t.getRow(32);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("ATD451601")) {
					rowt = sht_t.getRow(33);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("ATD451602")) {
					rowt = sht_t.getRow(34);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("ATD451603")) {
					rowt = sht_t.getRow(35);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("ATD4516R0")) {
					rowt = sht_t.getRow(36);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("-")) {
					rowt = sht_t.getRow(37);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("ATD451800")) {
					rowt = sht_t.getRow(38);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT0")) {
					rowt = sht_t.getRow(39);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT3")) {
					rowt = sht_t.getRow(40);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT6")) {
					rowt = sht_t.getRow(41);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT9")) {
					rowt = sht_t.getRow(42);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的华为天线型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(19).toString().equals("中兴")) {
				if (rowe.get(i).getCell(20).toString().equals("T-04-52-50-002")) {
					rowt = sht_t.getRow(43);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("T-03-52-52-003")) {
					rowt = sht_t.getRow(44);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("T-DA-02-00-59")) {
					rowt = sht_t.getRow(45);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("T-12-54-18-002")) {
					rowt = sht_t.getRow(46);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "第" + (i + 1) + "行" + "请输入正确的中兴天线型号！");
					System.exit(0);
				}
			} else if (rowe.get(i).getCell(19).toString().equals("大唐")
					|| rowe.get(i).getCell(19).toString().equals("上海贝尔")) {
				if (rowe.get(i).getCell(20).toString().equals("TYDA-202616D4T0")) {
					rowt = sht_t.getRow(47);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TYDA-202616D4T3")) {
					rowt = sht_t.getRow(48);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TYDA-202616D4T6")) {
					rowt = sht_t.getRow(49);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TYDA-202616D4T9")) {
					rowt = sht_t.getRow(50);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TYDA-2015/2616DE4-BC")) {
					rowt = sht_t.getRow(51);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TYDA-1917D4T0")) {
					rowt = sht_t.getRow(52);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TYDA-1917D4T3")) {
					rowt = sht_t.getRow(53);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT0")) {
					rowt = sht_t.getRow(54);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT3")) {
					rowt = sht_t.getRow(55);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT6")) {
					rowt = sht_t.getRow(56);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().equals("TDJ-172718D-65PT9")) {
					rowt = sht_t.getRow(57);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(rowt.getCell(j).toString());
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的大唐（或上海贝尔）天线型号！");
					System.exit(0);
				}
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线品牌！");
				System.exit(0);
			}
		}
		wbe.close();
		fise.close();
		wbt.close();
		fist.close();
		return tableMap;
	}
}
