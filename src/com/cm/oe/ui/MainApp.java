package com.cm.oe.ui;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JCheckBox;

import java.awt.BorderLayout;
import java.awt.Checkbox;
import java.awt.CheckboxGroup;
import java.awt.Component;
import java.awt.FlowLayout;
import java.awt.Label;
import java.awt.TextField;

import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;

import java.awt.GridLayout;
import java.awt.CardLayout;

import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.factories.FormFactory;
import com.jgoodies.forms.layout.RowSpec;

import javax.swing.JButton;
import javax.swing.JRadioButton;

import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.BoxLayout;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.ButtonGroup;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JLabel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.JCheckBoxMenuItem;
import javax.swing.JEditorPane;

import junit.framework.Test;

public class MainApp {

	private JFrame frame;

	/**
	 * Launch the application.
	 ****/
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainApp window = new MainApp();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public MainApp() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new MainFrame();
		frame.setBounds(100, 100, 565, 387);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		final TextField text=new TextField();
		text.setBounds(136, 246, 102, 15);
		text.setText("C:\\Users\\admin\\Desktop");
		frame.getContentPane().add(text);
		
		JButton jButton=new JButton();
		jButton.setBounds(262, 246, 72, 15);
		jButton.setText("...");
		jButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				        JFileChooser jFileChooser=new JFileChooser();
			            jFileChooser.setFileSelectionMode(1);  
			            int state = jFileChooser.showOpenDialog(null); 
			            if (state == 1) {  
			                return;  
			            } else {  
			                File file = jFileChooser.getSelectedFile(); 
			                text.setText(file.getAbsolutePath());  
			            }   
				
			}
		});
		frame.getContentPane().add(jButton);			
		
	    JLabel lblNewLabel = new JLabel("基站类型");
		lblNewLabel.setBounds(23, 28, 54, 15);
		frame.getContentPane().add(lblNewLabel);
		
		final JCheckBox a1 = new JCheckBox("宏基站");
		a1.setBounds(30, 54, 103, 23);
		//a1.addActionListener(al1);
		frame.getContentPane().add(a1);
		
		final JCheckBox a2 = new JCheckBox("小基站");
		a2.setBounds(162, 54, 103, 23);
		//a2.addActionListener(al2);
		frame.getContentPane().add(a2);
		
		final JCheckBox a3 = new JCheckBox("拉远站");
		a3.setBounds(277, 54, 103, 23);
		//a3.addActionListener(al3);
		frame.getContentPane().add(a3);
		
		final JCheckBox a4 = new JCheckBox("信源站");
		a4.setBounds(382, 54, 103, 23);
		//a4.addActionListener(al4);
		frame.getContentPane().add(a4);
		
		
		
		
		final JRadioButton b1 = new JRadioButton("室内站");
		b1.setBounds(30, 92, 76, 23);
		frame.getContentPane().add(b1);
		
		final JRadioButton b2 = new JRadioButton("室外站");
		b2.setBounds(162, 92, 76, 23);
		frame.getContentPane().add(b2);
		
		final ButtonGroup b = new ButtonGroup();
		b.add(b1);
		b.add(b2);
		
		
		JLabel lblNewLabel_1 = new JLabel("频段");
		lblNewLabel_1.setBounds(23, 121, 54, 15);
		frame.getContentPane().add(lblNewLabel_1);
		
		JRadioButton c1 = new JRadioButton("D频段");
		c1.setBounds(33, 145, 81, 23);
		frame.getContentPane().add(c1);
		
		JRadioButton c2 = new JRadioButton("E频段");
		c2.setBounds(162, 145, 81, 23);
		frame.getContentPane().add(c2);
		
		JRadioButton c3 = new JRadioButton("F频段");
		c3.setBounds(277, 145, 81, 23);
		frame.getContentPane().add(c3);
		
		ButtonGroup c = new ButtonGroup();
		c.add(c1);
		c.add(c2);
		c.add(c3);
		
		JLabel lblNewLabel_2 = new JLabel("生产厂家");
		lblNewLabel_2.setBounds(23, 184, 54, 15);
		frame.getContentPane().add(lblNewLabel_2);
		
		JRadioButton d1 = new JRadioButton("上海贝尔");
		d1.setBounds(33, 205, 81, 23);
		frame.getContentPane().add(d1);
		
		JRadioButton d2 = new JRadioButton("大唐");
		d2.setBounds(162, 205, 57, 23);
		frame.getContentPane().add(d2);
		
		JRadioButton d3 = new JRadioButton("华为");
		d3.setBounds(277, 205, 57, 23);
		frame.getContentPane().add(d3);
		
		JRadioButton d4 = new JRadioButton("中兴");
		d4.setBounds(382, 204, 57, 23);
		frame.getContentPane().add(d4);
		
		ButtonGroup d = new ButtonGroup();
		d.add(d1);
		d.add(d2);
		d.add(d3);
		d.add(d4);
		
		
		JButton btnNewButton = new JButton("导出表格");
		btnNewButton.setBounds(222, 286, 110, 23);
		btnNewButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				if (text.getText().equals("")||text.getText()==null) {
					JOptionPane.showMessageDialog(null, "请选择导出表格目录");
					return;					
				}else{
				int state = JOptionPane.showConfirmDialog(null, "确定导出?", "choose one", JOptionPane.YES_NO_OPTION);
				if(state==0){
			     @SuppressWarnings("resource")
				 HSSFWorkbook hssfWorkbook=new HSSFWorkbook();
			     HSSFSheet hssfSheet=hssfWorkbook.createSheet();
			     HSSFCellStyle style = hssfWorkbook.createCellStyle();
			     HSSFFont  font =hssfWorkbook.createFont();
			     font.setFontName("宋体");
			     font.setFontHeightInPoints((short) 14);
			     style.setFont(font);
			     
			     HSSFRow hssfRow=hssfSheet.createRow(0);			 
			     HSSFRow hssfRow1=hssfSheet.createRow(1);
			     
			     HSSFCell hssfCell=hssfRow.createCell(0);
			     hssfCell.setCellValue("归属地市   ");
			     
			     HSSFCell hssfCell1=hssfRow.createCell(1);
			     hssfCell1.setCellValue("站名（站号）  ");
			     
			     HSSFCell hssfCell2=hssfRow.createCell(2);
			     hssfCell2.setCellValue("第几册 " );
			     
			     HSSFCell hssfCell3=hssfRow.createCell(3);
			     hssfCell3.setCellValue("地市设计编号   ");
			     
			     HSSFCell hssfCell4=hssfRow.createCell(4);
			     hssfCell4.setCellValue("设计完成月份   ");
			     
			     HSSFCell hssfCell5=hssfRow.createCell(5);
			     hssfCell5.setCellValue("专业审核人   ");
			     
			     HSSFCell hssfCell6=hssfRow.createCell(6);
			     hssfCell6.setCellValue("单项负责人   ");
			     
			     HSSFCell hssfCell7=hssfRow.createCell(7);
			     hssfCell7.setCellValue("概预算审核人   ");
			     
			     HSSFCell hssfCell8=hssfRow.createCell(8);
			     hssfCell8.setCellValue("概预算编制人   ");
			     
			     HSSFCell hssfCell9=hssfRow.createCell(9);
			     hssfCell9.setCellValue("详细站址   ");
			     
			     HSSFCell hssfCell10=hssfRow.createCell(10);
			     hssfCell10.setCellValue("覆盖区域名称   ");
			     
			     HSSFCell hssfCell11=hssfRow.createCell(11);
			     hssfCell11.setCellValue("经度");
			     
			     HSSFCell hssfCell12=hssfRow.createCell(12);
			     hssfCell12.setCellValue("纬度 ");
			     
			     HSSFCell hssfCell13=hssfRow.createCell(13);
			     hssfCell13.setCellValue("本工程建设规模   ");
			     
			     HSSFCell hssfCell14=hssfRow.createCell(14);
			     hssfCell14.setCellValue("预立项文件   ");
			     
			     HSSFCell hssfCell15=hssfRow.createCell(15);
			     hssfCell15.setCellValue("BBU设备  ");
			     
			     HSSFCell hssfCell16=hssfRow.createCell(16);
			     hssfCell16.setCellValue("RRU设备  ");
			     
			     HSSFCell hssfCell17=hssfRow.createCell(17);
			     hssfCell17.setCellValue("天线设备  ");
			     
			     HSSFCell hssfCell18=hssfRow.createCell(18);
			     hssfCell18.setCellValue("抗震设防烈度   ");
			     
			     HSSFCell hssfCell19=hssfRow.createCell(19);
			     hssfCell19.setCellValue("主设备安装方式   ");
			     
			     HSSFCell hssfCell20=hssfRow.createCell(20);
			     hssfCell20.setCellValue("天线方位角   ");
			     
			     HSSFCell hssfCell21=hssfRow.createCell(21);
			     hssfCell21.setCellValue("天线挂高   ");
			     
			     HSSFCell hssfCell22=hssfRow.createCell(22);
			     hssfCell22.setCellValue("总下倾角  ");
			     
			     HSSFCell hssfCell23=hssfRow.createCell(23);
			     hssfCell23.setCellValue("天馈情况   ");
			     
			     HSSFCell hssfCell24=hssfRow.createCell(24);
			     hssfCell24.setCellValue("配置   ");
			     
			     HSSFCell hssfCell25=hssfRow.createCell(25);
			     hssfCell25.setCellValue("RRU数量   ");
			     
			     HSSFCell hssfCell26=hssfRow.createCell(26);
			     hssfCell26.setCellValue("现网覆盖状况以及存在问题      ");
			     int i;
			     for(i=0;i<27;i++){
			    	 HSSFCell hf = hssfRow.getCell(i);
			    	 hf.setCellStyle(style);
			     }
			     
			     hssfSheet.autoSizeColumn(0);
			     hssfSheet.autoSizeColumn(1);
			     hssfSheet.autoSizeColumn(2);
			     hssfSheet.autoSizeColumn(3);
			     hssfSheet.autoSizeColumn(4);
			     hssfSheet.autoSizeColumn(5);
			     hssfSheet.autoSizeColumn(6);
			     hssfSheet.autoSizeColumn(7);
			     hssfSheet.autoSizeColumn(8);
			     hssfSheet.autoSizeColumn(9);
			     hssfSheet.autoSizeColumn(10);
			     hssfSheet.autoSizeColumn(11);
			     hssfSheet.autoSizeColumn(12);
			     hssfSheet.autoSizeColumn(13);
			     hssfSheet.autoSizeColumn(14);
			     hssfSheet.autoSizeColumn(15);
			     hssfSheet.autoSizeColumn(16);
			     hssfSheet.autoSizeColumn(17);
			     hssfSheet.autoSizeColumn(18);
			     hssfSheet.autoSizeColumn(19);
			     hssfSheet.autoSizeColumn(20);
			     hssfSheet.autoSizeColumn(21);
			     hssfSheet.autoSizeColumn(22);
			     hssfSheet.autoSizeColumn(23);
			     hssfSheet.autoSizeColumn(24);
			     hssfSheet.autoSizeColumn(25);
			     hssfSheet.autoSizeColumn(26); 
			     
			 	if (a1.getSelectedObjects()!=null) {
			 		 HSSFCell hssfCell100=hssfRow1.createCell(0);
				     hssfCell100.setCellValue(a1.getText());
				}
			
			 	if (a2.getSelectedObjects()!=null){
			     HSSFCell hssfCell101=hssfRow1.createCell(1);
			     hssfCell101.setCellValue(a2.getText());
			 	}
			     
			     
			     
			     try{
			     FileOutputStream fileOutputStream=new FileOutputStream(text.getText()+"/"+Math.round(Math.random()*1000000)+".xls");
			     hssfWorkbook.write(fileOutputStream);
			     fileOutputStream.flush();
			     fileOutputStream.close();
			     }catch(IOException e1){
			    	 e1.printStackTrace();
			    	 
			     }
			  JOptionPane.showMessageDialog(null, "导出完成");
				}else{
					return;
				}}
			}
		});
		frame.getContentPane().add(btnNewButton);
	
		JLabel lblNewLabel_3 = new JLabel("请选择保存地址");
		lblNewLabel_3.setBounds(23, 246, 107, 15);
		frame.getContentPane().add(lblNewLabel_3);
		
		
	
	}
}
