package com.cm.oe.test;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class ReadWord {
	public Range getRange(HWPFDocument doc) {
		Range range = doc.getRange();
		return range;
	}

	public List<String> getParavalue(Range range) {
		List<String> results = new ArrayList<String>();
		Matcher mat = matcherPara(range.text());
		while (mat.find()) {
			results.add(mat.group());
		}
		return results;
	}
	
	public List<String> getTablevalue(Range range){
		List<String> results = new ArrayList<String>();
		Matcher mat = matcherTable(range.text());
		while(mat.find()){
			results.add(mat.group());
		}
		return results;
	}
	
	public List<String> getOptionvalue(Range range){
		List<String> results = new ArrayList<String>();
		Matcher mat = matcherTable(range.text());
		while(mat.find()){
			results.add(mat.group());
		}
		return results;
	}
	
	public Matcher matcherPara(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
	public Matcher matcherTable(String str){
		Pattern pattern = Pattern.compile("\\#\\{(.+?)\\}",Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
	public Matcher matcherOptionFile(String str){
		Pattern pattern = Pattern.compile("\\@\\{(.+?)\\}",Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
}
