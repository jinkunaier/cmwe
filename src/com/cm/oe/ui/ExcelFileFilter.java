package com.cm.oe.ui;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class ExcelFileFilter extends FileFilter {

	@Override
	public boolean accept(File f) {
		// TODO Auto-generated method stub
		String name = f.getName();    
        return f.isDirectory() || name.toLowerCase().endsWith(".xls") ;
	}

	@Override
	public String getDescription() {
		// TODO Auto-generated method stub
		return "*.xls";
	}

}
