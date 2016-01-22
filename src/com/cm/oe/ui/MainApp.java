package com.cm.oe.ui;

import java.awt.EventQueue;
import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JRadioButton;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JPanel;
import java.awt.GridLayout;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import java.awt.Color;
import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.Insets;

public class MainApp {
	
	String aText;
	String bText;
	String cText;
	String dText;


	private JFrame frame;

	/**
	 * Launch the application.
	 ****/


	/**
	 * Create the application.
	 */
	public MainApp() {
		initialize();
		this.frame.setVisible(true);
		
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new MainFrame();
		frame.setResizable(false);
		frame.setTitle("导出Excel");
		frame.setBounds(100, 100, 711, 438);
		frame.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
		/*给四组按钮添加事件*/
		ActionListener aListener =new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				aText=((JRadioButton)e.getSource()).getText();
				
			}
		};
		
        ActionListener bListener =new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				bText=((JRadioButton)e.getSource()).getText();
				
			}
		};
        ActionListener cListener =new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				cText=((JRadioButton)e.getSource()).getText();	
			}
		};
		
       ActionListener dListener =new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				dText=((JRadioButton)e.getSource()).getText();	
			}
		};
		
		final TextField text=new TextField();
		
		final ButtonGroup a = new ButtonGroup();
		
		final ButtonGroup b = new ButtonGroup();
		
		final ButtonGroup c = new ButtonGroup();
		
		final ButtonGroup d = new ButtonGroup();
		GridBagLayout gridBagLayout = new GridBagLayout();
		gridBagLayout.columnWidths = new int[]{118, 118, 118, 118, 118, 118, 0};
		gridBagLayout.rowHeights = new int[]{43, 43, 43, 43, 43, 43, 43, 43, 43, 0};
		gridBagLayout.columnWeights = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		gridBagLayout.rowWeights = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		frame.getContentPane().setLayout(gridBagLayout);
		
		JRadioButton d3 = new JRadioButton("华为");
		d3.addActionListener(dListener);
		
		
		
		
		
		JRadioButton b1 = new JRadioButton("室内站");
		b1.addActionListener(bListener);
		
		JRadioButton a1 = new JRadioButton("宏基站");
		a1.addActionListener(aListener);
		
		JLabel label = new JLabel("");
		GridBagConstraints gbc_label = new GridBagConstraints();
		gbc_label.fill = GridBagConstraints.BOTH;
		gbc_label.insets = new Insets(0, 0, 5, 5);
		gbc_label.gridx = 0;
		gbc_label.gridy = 0;
		frame.getContentPane().add(label, gbc_label);
		JLabel lblNewLabel = new JLabel("基站类型");
		GridBagConstraints gbc_lblNewLabel = new GridBagConstraints();
		gbc_lblNewLabel.fill = GridBagConstraints.BOTH;
		gbc_lblNewLabel.insets = new Insets(0, 0, 5, 5);
		gbc_lblNewLabel.gridx = 1;
		gbc_lblNewLabel.gridy = 0;
		frame.getContentPane().add(lblNewLabel, gbc_lblNewLabel);
		
		JLabel label_1 = new JLabel("");
		GridBagConstraints gbc_label_1 = new GridBagConstraints();
		gbc_label_1.fill = GridBagConstraints.BOTH;
		gbc_label_1.insets = new Insets(0, 0, 5, 5);
		gbc_label_1.gridx = 2;
		gbc_label_1.gridy = 0;
		frame.getContentPane().add(label_1, gbc_label_1);
		
		JLabel label_2 = new JLabel("");
		GridBagConstraints gbc_label_2 = new GridBagConstraints();
		gbc_label_2.fill = GridBagConstraints.BOTH;
		gbc_label_2.insets = new Insets(0, 0, 5, 5);
		gbc_label_2.gridx = 3;
		gbc_label_2.gridy = 0;
		frame.getContentPane().add(label_2, gbc_label_2);
		
		JLabel label_3 = new JLabel("");
		GridBagConstraints gbc_label_3 = new GridBagConstraints();
		gbc_label_3.fill = GridBagConstraints.BOTH;
		gbc_label_3.insets = new Insets(0, 0, 5, 5);
		gbc_label_3.gridx = 4;
		gbc_label_3.gridy = 0;
		frame.getContentPane().add(label_3, gbc_label_3);
		
		JLabel label_4 = new JLabel("");
		GridBagConstraints gbc_label_4 = new GridBagConstraints();
		gbc_label_4.fill = GridBagConstraints.BOTH;
		gbc_label_4.insets = new Insets(0, 0, 5, 0);
		gbc_label_4.gridx = 5;
		gbc_label_4.gridy = 0;
		frame.getContentPane().add(label_4, gbc_label_4);
		
		JLabel label_5 = new JLabel("");
		GridBagConstraints gbc_label_5 = new GridBagConstraints();
		gbc_label_5.fill = GridBagConstraints.BOTH;
		gbc_label_5.insets = new Insets(0, 0, 5, 5);
		gbc_label_5.gridx = 0;
		gbc_label_5.gridy = 1;
		frame.getContentPane().add(label_5, gbc_label_5);
		GridBagConstraints gbc_a1 = new GridBagConstraints();
		gbc_a1.fill = GridBagConstraints.VERTICAL;
		gbc_a1.insets = new Insets(0, 0, 5, 5);
		gbc_a1.gridx = 1;
		gbc_a1.gridy = 1;
		frame.getContentPane().add(a1, gbc_a1);
		a.add(a1);
		
		JRadioButton a3 = new JRadioButton("拉远站");
		a3.addActionListener(aListener);
		
		JRadioButton a2 = new JRadioButton("小基站");
		a2.addActionListener(aListener);
		GridBagConstraints gbc_a2 = new GridBagConstraints();
		gbc_a2.fill = GridBagConstraints.VERTICAL;
		gbc_a2.insets = new Insets(0, 0, 5, 5);
		gbc_a2.gridx = 2;
		gbc_a2.gridy = 1;
		frame.getContentPane().add(a2, gbc_a2);
		a.add(a2);
		GridBagConstraints gbc_a3 = new GridBagConstraints();
		gbc_a3.fill = GridBagConstraints.VERTICAL;
		gbc_a3.insets = new Insets(0, 0, 5, 5);
		gbc_a3.gridx = 3;
		gbc_a3.gridy = 1;
		frame.getContentPane().add(a3, gbc_a3);
		a.add(a3);
		
		JRadioButton a4 = new JRadioButton("信源站");
		a4.addActionListener(aListener);
		GridBagConstraints gbc_a4 = new GridBagConstraints();
		gbc_a4.fill = GridBagConstraints.VERTICAL;
		gbc_a4.insets = new Insets(0, 0, 5, 5);
		gbc_a4.gridx = 4;
		gbc_a4.gridy = 1;
		frame.getContentPane().add(a4, gbc_a4);
		a.add(a4);
		
		JLabel label_6 = new JLabel("");
		GridBagConstraints gbc_label_6 = new GridBagConstraints();
		gbc_label_6.fill = GridBagConstraints.BOTH;
		gbc_label_6.insets = new Insets(0, 0, 5, 0);
		gbc_label_6.gridx = 5;
		gbc_label_6.gridy = 1;
		frame.getContentPane().add(label_6, gbc_label_6);
		
		JLabel label_7 = new JLabel("");
		GridBagConstraints gbc_label_7 = new GridBagConstraints();
		gbc_label_7.fill = GridBagConstraints.BOTH;
		gbc_label_7.insets = new Insets(0, 0, 5, 5);
		gbc_label_7.gridx = 0;
		gbc_label_7.gridy = 2;
		frame.getContentPane().add(label_7, gbc_label_7);
		b.add(b1);
		GridBagConstraints gbc_b1 = new GridBagConstraints();
		gbc_b1.fill = GridBagConstraints.VERTICAL;
		gbc_b1.insets = new Insets(0, 0, 5, 5);
		gbc_b1.gridx = 1;
		gbc_b1.gridy = 2;
		frame.getContentPane().add(b1, gbc_b1);
		final JRadioButton b2 = new JRadioButton("室外站");
		b2.addActionListener(bListener);
		b.add(b2);
		GridBagConstraints gbc_b2 = new GridBagConstraints();
		gbc_b2.fill = GridBagConstraints.VERTICAL;
		gbc_b2.insets = new Insets(0, 0, 5, 5);
		gbc_b2.gridx = 2;
		gbc_b2.gridy = 2;
		frame.getContentPane().add(b2, gbc_b2);
		
		JLabel label_8 = new JLabel("");
		GridBagConstraints gbc_label_8 = new GridBagConstraints();
		gbc_label_8.fill = GridBagConstraints.BOTH;
		gbc_label_8.insets = new Insets(0, 0, 5, 5);
		gbc_label_8.gridx = 3;
		gbc_label_8.gridy = 2;
		frame.getContentPane().add(label_8, gbc_label_8);
		
		JLabel label_9 = new JLabel("");
		GridBagConstraints gbc_label_9 = new GridBagConstraints();
		gbc_label_9.fill = GridBagConstraints.BOTH;
		gbc_label_9.insets = new Insets(0, 0, 5, 5);
		gbc_label_9.gridx = 4;
		gbc_label_9.gridy = 2;
		frame.getContentPane().add(label_9, gbc_label_9);
		
		JLabel label_10 = new JLabel("");
		GridBagConstraints gbc_label_10 = new GridBagConstraints();
		gbc_label_10.fill = GridBagConstraints.BOTH;
		gbc_label_10.insets = new Insets(0, 0, 5, 0);
		gbc_label_10.gridx = 5;
		gbc_label_10.gridy = 2;
		frame.getContentPane().add(label_10, gbc_label_10);
		
		JLabel label_11 = new JLabel("");
		GridBagConstraints gbc_label_11 = new GridBagConstraints();
		gbc_label_11.fill = GridBagConstraints.BOTH;
		gbc_label_11.insets = new Insets(0, 0, 5, 5);
		gbc_label_11.gridx = 0;
		gbc_label_11.gridy = 3;
		frame.getContentPane().add(label_11, gbc_label_11);
		
		
		JLabel lblNewLabel_1 = new JLabel("频段");
		GridBagConstraints gbc_lblNewLabel_1 = new GridBagConstraints();
		gbc_lblNewLabel_1.fill = GridBagConstraints.BOTH;
		gbc_lblNewLabel_1.insets = new Insets(0, 0, 5, 5);
		gbc_lblNewLabel_1.gridx = 1;
		gbc_lblNewLabel_1.gridy = 3;
		frame.getContentPane().add(lblNewLabel_1, gbc_lblNewLabel_1);
		
		JLabel label_12 = new JLabel("");
		GridBagConstraints gbc_label_12 = new GridBagConstraints();
		gbc_label_12.fill = GridBagConstraints.BOTH;
		gbc_label_12.insets = new Insets(0, 0, 5, 5);
		gbc_label_12.gridx = 2;
		gbc_label_12.gridy = 3;
		frame.getContentPane().add(label_12, gbc_label_12);
		
		JLabel label_13 = new JLabel("");
		GridBagConstraints gbc_label_13 = new GridBagConstraints();
		gbc_label_13.fill = GridBagConstraints.BOTH;
		gbc_label_13.insets = new Insets(0, 0, 5, 5);
		gbc_label_13.gridx = 3;
		gbc_label_13.gridy = 3;
		frame.getContentPane().add(label_13, gbc_label_13);
		
		JRadioButton c2 = new JRadioButton("E频段");
		c2.addActionListener(cListener);
		
		JRadioButton c1 = new JRadioButton("D频段");
		c1.addActionListener(cListener);
		
		JLabel label_14 = new JLabel("");
		GridBagConstraints gbc_label_14 = new GridBagConstraints();
		gbc_label_14.fill = GridBagConstraints.BOTH;
		gbc_label_14.insets = new Insets(0, 0, 5, 5);
		gbc_label_14.gridx = 4;
		gbc_label_14.gridy = 3;
		frame.getContentPane().add(label_14, gbc_label_14);
		
		JLabel label_15 = new JLabel("");
		GridBagConstraints gbc_label_15 = new GridBagConstraints();
		gbc_label_15.fill = GridBagConstraints.BOTH;
		gbc_label_15.insets = new Insets(0, 0, 5, 0);
		gbc_label_15.gridx = 5;
		gbc_label_15.gridy = 3;
		frame.getContentPane().add(label_15, gbc_label_15);
		
		JLabel label_16 = new JLabel("");
		GridBagConstraints gbc_label_16 = new GridBagConstraints();
		gbc_label_16.fill = GridBagConstraints.BOTH;
		gbc_label_16.insets = new Insets(0, 0, 5, 5);
		gbc_label_16.gridx = 0;
		gbc_label_16.gridy = 4;
		frame.getContentPane().add(label_16, gbc_label_16);
		c.add(c1);
		GridBagConstraints gbc_c1 = new GridBagConstraints();
		gbc_c1.fill = GridBagConstraints.VERTICAL;
		gbc_c1.insets = new Insets(0, 0, 5, 5);
		gbc_c1.gridx = 1;
		gbc_c1.gridy = 4;
		frame.getContentPane().add(c1, gbc_c1);
		c.add(c2);
		GridBagConstraints gbc_c2 = new GridBagConstraints();
		gbc_c2.fill = GridBagConstraints.VERTICAL;
		gbc_c2.insets = new Insets(0, 0, 5, 5);
		gbc_c2.gridx = 2;
		gbc_c2.gridy = 4;
		frame.getContentPane().add(c2, gbc_c2);
		
		JRadioButton c3 = new JRadioButton("F频段");
		c3.addActionListener(cListener);
		c.add(c3);
		GridBagConstraints gbc_c3 = new GridBagConstraints();
		gbc_c3.fill = GridBagConstraints.VERTICAL;
		gbc_c3.insets = new Insets(0, 0, 5, 5);
		gbc_c3.gridx = 3;
		gbc_c3.gridy = 4;
		frame.getContentPane().add(c3, gbc_c3);
		
		JLabel label_17 = new JLabel("");
		GridBagConstraints gbc_label_17 = new GridBagConstraints();
		gbc_label_17.fill = GridBagConstraints.BOTH;
		gbc_label_17.insets = new Insets(0, 0, 5, 5);
		gbc_label_17.gridx = 4;
		gbc_label_17.gridy = 4;
		frame.getContentPane().add(label_17, gbc_label_17);
		
		JLabel label_18 = new JLabel("");
		GridBagConstraints gbc_label_18 = new GridBagConstraints();
		gbc_label_18.fill = GridBagConstraints.BOTH;
		gbc_label_18.insets = new Insets(0, 0, 5, 0);
		gbc_label_18.gridx = 5;
		gbc_label_18.gridy = 4;
		frame.getContentPane().add(label_18, gbc_label_18);
		
		JLabel label_19 = new JLabel("");
		GridBagConstraints gbc_label_19 = new GridBagConstraints();
		gbc_label_19.fill = GridBagConstraints.BOTH;
		gbc_label_19.insets = new Insets(0, 0, 5, 5);
		gbc_label_19.gridx = 0;
		gbc_label_19.gridy = 5;
		frame.getContentPane().add(label_19, gbc_label_19);
		
		JLabel lblNewLabel_2 = new JLabel("生产厂家");
		GridBagConstraints gbc_lblNewLabel_2 = new GridBagConstraints();
		gbc_lblNewLabel_2.fill = GridBagConstraints.BOTH;
		gbc_lblNewLabel_2.insets = new Insets(0, 0, 5, 5);
		gbc_lblNewLabel_2.gridx = 1;
		gbc_lblNewLabel_2.gridy = 5;
		frame.getContentPane().add(lblNewLabel_2, gbc_lblNewLabel_2);
		
		JLabel label_20 = new JLabel("");
		GridBagConstraints gbc_label_20 = new GridBagConstraints();
		gbc_label_20.fill = GridBagConstraints.BOTH;
		gbc_label_20.insets = new Insets(0, 0, 5, 5);
		gbc_label_20.gridx = 2;
		gbc_label_20.gridy = 5;
		frame.getContentPane().add(label_20, gbc_label_20);
		
		JLabel label_21 = new JLabel("");
		GridBagConstraints gbc_label_21 = new GridBagConstraints();
		gbc_label_21.fill = GridBagConstraints.BOTH;
		gbc_label_21.insets = new Insets(0, 0, 5, 5);
		gbc_label_21.gridx = 3;
		gbc_label_21.gridy = 5;
		frame.getContentPane().add(label_21, gbc_label_21);
		
		JLabel label_22 = new JLabel("");
		GridBagConstraints gbc_label_22 = new GridBagConstraints();
		gbc_label_22.fill = GridBagConstraints.BOTH;
		gbc_label_22.insets = new Insets(0, 0, 5, 5);
		gbc_label_22.gridx = 4;
		gbc_label_22.gridy = 5;
		frame.getContentPane().add(label_22, gbc_label_22);
		
		JLabel label_23 = new JLabel("");
		GridBagConstraints gbc_label_23 = new GridBagConstraints();
		gbc_label_23.fill = GridBagConstraints.BOTH;
		gbc_label_23.insets = new Insets(0, 0, 5, 0);
		gbc_label_23.gridx = 5;
		gbc_label_23.gridy = 5;
		frame.getContentPane().add(label_23, gbc_label_23);
		
		JRadioButton d1 = new JRadioButton("上海贝尔");
		d1.addActionListener(dListener);
		
		JLabel label_24 = new JLabel("");
		GridBagConstraints gbc_label_24 = new GridBagConstraints();
		gbc_label_24.fill = GridBagConstraints.BOTH;
		gbc_label_24.insets = new Insets(0, 0, 5, 5);
		gbc_label_24.gridx = 0;
		gbc_label_24.gridy = 6;
		frame.getContentPane().add(label_24, gbc_label_24);
		d.add(d1);
		GridBagConstraints gbc_d1 = new GridBagConstraints();
		gbc_d1.fill = GridBagConstraints.VERTICAL;
		gbc_d1.insets = new Insets(0, 0, 5, 5);
		gbc_d1.gridx = 1;
		gbc_d1.gridy = 6;
		frame.getContentPane().add(d1, gbc_d1);
		
		JRadioButton d2 = new JRadioButton("大唐");
		d2.addActionListener(dListener);
		d.add(d2);
		GridBagConstraints gbc_d2 = new GridBagConstraints();
		gbc_d2.fill = GridBagConstraints.VERTICAL;
		gbc_d2.insets = new Insets(0, 0, 5, 5);
		gbc_d2.gridx = 2;
		gbc_d2.gridy = 6;
		frame.getContentPane().add(d2, gbc_d2);
		d.add(d3);
		GridBagConstraints gbc_d3 = new GridBagConstraints();
		gbc_d3.fill = GridBagConstraints.VERTICAL;
		gbc_d3.insets = new Insets(0, 0, 5, 5);
		gbc_d3.gridx = 3;
		gbc_d3.gridy = 6;
		frame.getContentPane().add(d3, gbc_d3);
		
		//为导出表格添加事件
		JButton btnNewButton = new JButton("导出表格");
		btnNewButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				if (text.getText().equals("")||text.getText()==null) {
					JOptionPane.showMessageDialog(null, "请选择导出表格目录");
					return;					
				}else{
				if (a.getSelection()==null||b.getSelection()==null||c .getSelection()==null||d.getSelection()==null) {

					     JOptionPane.showMessageDialog(null, "请选择全部条件");
					     return;
					}
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
			     if(aText.equals("信源站")){
			    	 HSSFCell hssfCell=hssfRow.createCell(0);
				     hssfCell.setCellValue("归属地市   ");
				     
				     HSSFCell hssfCell1=hssfRow.createCell(1);
				     hssfCell1.setCellValue("站名  ");
				     
				     HSSFCell hssfCell2=hssfRow.createCell(2);
				     hssfCell2.setCellValue("站号  ");
				     
				     HSSFCell hssfCell3=hssfRow.createCell(3);
				     hssfCell3.setCellValue("第几册 " );
				     
				     HSSFCell hssfCell4=hssfRow.createCell(4);
				     hssfCell4.setCellValue("地市设计编号   ");
				     
				     HSSFCell hssfCell5=hssfRow.createCell(5);
				     hssfCell5.setCellValue("设计完成月份   ");
				     
				     HSSFCell hssfCell6=hssfRow.createCell(6);
				     hssfCell6.setCellValue("专业审核人   ");
				     
				     HSSFCell hssfCell7=hssfRow.createCell(7);
				     hssfCell7.setCellValue("单项负责人   ");
				     
				     HSSFCell hssfCell8=hssfRow.createCell(8);
				     hssfCell8.setCellValue("概预算审核人   ");
				     
				     HSSFCell hssfCell9=hssfRow.createCell(9);
				     hssfCell9.setCellValue("概预算编制人   ");
				     
				     HSSFCell hssfCell10=hssfRow.createCell(10);
				     hssfCell10.setCellValue("详细站址   ");
				     
				     HSSFCell hssfCell11=hssfRow.createCell(11);
				     hssfCell11.setCellValue("覆盖区域名称   ");
				     
				     HSSFCell hssfCell12=hssfRow.createCell(12);
				     hssfCell12.setCellValue("经度 ");
				     
				     HSSFCell hssfCell13=hssfRow.createCell(13);
				     hssfCell13.setCellValue("纬度 ");
				     
				     HSSFCell hssfCell14=hssfRow.createCell(14);
				     hssfCell14.setCellValue("宏基站   ");
				     
				     HSSFCell hssfCell15=hssfRow.createCell(15);
				     hssfCell15.setCellValue("小基站   ");
				     
				     HSSFCell hssfCell16=hssfRow.createCell(16);
				     hssfCell16.setCellValue("拉远站   ");
				     
				     HSSFCell hssfCell17=hssfRow.createCell(17);
				     hssfCell17.setCellValue("信源站   ");
				     
				     HSSFCell hssfCell18=hssfRow.createCell(18);
				     hssfCell18.setCellValue("预立项文件  ");
				     
				     HSSFCell hssfCell19=hssfRow.createCell(19);
				     hssfCell19.setCellValue("BBU品牌  ");
				     
				     HSSFCell hssfCell20=hssfRow.createCell(20);
				     hssfCell20.setCellValue("BBU型号  ");
				     
				     HSSFCell hssfCell21=hssfRow.createCell(21);
				     hssfCell21.setCellValue("RRU品牌  ");
				     
				     HSSFCell hssfCell22=hssfRow.createCell(22);
				     hssfCell22.setCellValue("RRU型号  ");
				     
				     HSSFCell hssfCell23=hssfRow.createCell(23);
				     hssfCell23.setCellValue("天线品牌 ");
				     
				     HSSFCell hssfCell24=hssfRow.createCell(24);
				     hssfCell24.setCellValue("天线型号 ");
				     
				     HSSFCell hssfCell25=hssfRow.createCell(25);
				     hssfCell25.setCellValue("抗震设防烈度   ");
				     
				     HSSFCell hssfCell26=hssfRow.createCell(26);
				     hssfCell26.setCellValue("主设备安装方式   ");
				     
				     HSSFCell hssfCell27=hssfRow.createCell(27);
				     hssfCell27.setCellValue("工程类型   ");
				     
				     HSSFCell hssfCell28=hssfRow.createCell(28);
				     hssfCell28.setCellValue("配置   ");
				     
				     HSSFCell hssfCell29=hssfRow.createCell(29);
				     hssfCell29.setCellValue("RRU数量      ");
				     
				     HSSFCell hssfCell30=hssfRow.createCell(30);
				     hssfCell30.setCellValue("现网覆盖状况以及存在问题      ");
				     
				     int i;
				     for(i=0;i<31;i++){
				    	 HSSFCell hf = hssfRow.getCell(i);
				    	 hf.setCellStyle(style);
				    	 hssfSheet.autoSizeColumn(i);
				     }				   			    	 
			     }
			     			 
			     if(aText.equals("拉远站"))
			     {
			     HSSFCell hssfCell=hssfRow.createCell(0);
			     hssfCell.setCellValue("归属地市   ");
			     
			     HSSFCell hssfCell1=hssfRow.createCell(1);
			     hssfCell1.setCellValue("站名  ");
			     
			     HSSFCell hssfCell2=hssfRow.createCell(2);
			     hssfCell2.setCellValue("站号  ");
			     
			     HSSFCell hssfCell3=hssfRow.createCell(3);
			     hssfCell3.setCellValue("第几册 " );
			     
			     HSSFCell hssfCell4=hssfRow.createCell(4);
			     hssfCell4.setCellValue("地市设计编号   ");
			     
			     HSSFCell hssfCell5=hssfRow.createCell(5);
			     hssfCell5.setCellValue("设计完成月份   ");
			     
			     HSSFCell hssfCell6=hssfRow.createCell(6);
			     hssfCell6.setCellValue("专业审核人   ");
			     
			     HSSFCell hssfCell7=hssfRow.createCell(7);
			     hssfCell7.setCellValue("单项负责人   ");
			     
			     HSSFCell hssfCell8=hssfRow.createCell(8);
			     hssfCell8.setCellValue("概预算审核人   ");
			     
			     HSSFCell hssfCell9=hssfRow.createCell(9);
			     hssfCell9.setCellValue("概预算编制人   ");
			     
			     HSSFCell hssfCell10=hssfRow.createCell(10);
			     hssfCell10.setCellValue("详细站址   ");
			     
			     HSSFCell hssfCell11=hssfRow.createCell(11);
			     hssfCell11.setCellValue("覆盖区域名称   ");
			     
			     HSSFCell hssfCell12=hssfRow.createCell(12);
			     hssfCell12.setCellValue("经度 ");
			     
			     HSSFCell hssfCell13=hssfRow.createCell(13);
			     hssfCell13.setCellValue("纬度 ");
			     
			     HSSFCell hssfCell14=hssfRow.createCell(14);
			     hssfCell14.setCellValue("宏基站   ");
			     
			     HSSFCell hssfCell15=hssfRow.createCell(15);
			     hssfCell15.setCellValue("小基站   ");
			     
			     HSSFCell hssfCell16=hssfRow.createCell(16);
			     hssfCell16.setCellValue("拉远站   ");
			     
			     HSSFCell hssfCell17=hssfRow.createCell(17);
			     hssfCell17.setCellValue("信源站   ");
			     
			     HSSFCell hssfCell18=hssfRow.createCell(18);
			     hssfCell18.setCellValue("预立项文件  ");
			     
			     HSSFCell hssfCell19=hssfRow.createCell(19);
			     hssfCell19.setCellValue("BBU品牌  ");
			     
			     HSSFCell hssfCell20=hssfRow.createCell(20);
			     hssfCell20.setCellValue("BBU型号  ");
			     
			     HSSFCell hssfCell21=hssfRow.createCell(21);
			     hssfCell21.setCellValue("RRU品牌  ");
			     
			     HSSFCell hssfCell22=hssfRow.createCell(22);
			     hssfCell22.setCellValue("RRU型号  ");
			     
			     HSSFCell hssfCell23=hssfRow.createCell(23);
			     hssfCell23.setCellValue("天线品牌 ");
			     
			     HSSFCell hssfCell24=hssfRow.createCell(24);
			     hssfCell24.setCellValue("天线型号 ");
			     
			     HSSFCell hssfCell25=hssfRow.createCell(25);
			     hssfCell25.setCellValue("抗震设防烈度   ");
			     
			     HSSFCell hssfCell26=hssfRow.createCell(26);
			     hssfCell26.setCellValue("主设备安装方式   ");
			     
			     HSSFCell hssfCell27=hssfRow.createCell(27);
			     hssfCell27.setCellValue("工程类型   ");
			  
			     HSSFCell hssfCell28=hssfRow.createCell(28);
			     hssfCell28.setCellValue("天线方位角   ");
			     
			     HSSFCell hssfCell29=hssfRow.createCell(29);
			     hssfCell29.setCellValue("天线挂高  ");
			     
			     HSSFCell hssfCell30=hssfRow.createCell(30);
			     hssfCell30.setCellValue("总下倾角   ");
			     
			     HSSFCell hssfCell31=hssfRow.createCell(31);
			     hssfCell31.setCellValue("天馈情况   ");
			     
			     HSSFCell hssfCell32=hssfRow.createCell(32);
			     hssfCell32.setCellValue("配置   ");
			     
			     HSSFCell hssfCell33=hssfRow.createCell(33);
			     hssfCell33.setCellValue("RRU数量      ");
			     
			     HSSFCell hssfCell34=hssfRow.createCell(34);
			     hssfCell34.setCellValue("现网覆盖状况以及存在问题      ");
			     
			     int i;
			     for(i=0;i<35;i++){
			    	 HSSFCell hf = hssfRow.getCell(i);
			    	 hf.setCellStyle(style);
			    	 hssfSheet.autoSizeColumn(i);
			     }
			     
			     }
			 
			     if(aText.equals("宏基站")||aText.equals("小基站")){
			    	 HSSFCell hssfCell=hssfRow.createCell(0);
				     hssfCell.setCellValue("归属地市   ");
				     
				     HSSFCell hssfCell1=hssfRow.createCell(1);
				     hssfCell1.setCellValue("站名  ");
				     
				     HSSFCell hssfCell2=hssfRow.createCell(2);
				     hssfCell2.setCellValue("站号  ");
				     
				     HSSFCell hssfCell3=hssfRow.createCell(3);
				     hssfCell3.setCellValue("第几册 " );
				     
				     HSSFCell hssfCell4=hssfRow.createCell(4);
				     hssfCell4.setCellValue("地市设计编号   ");
				     
				     HSSFCell hssfCell5=hssfRow.createCell(5);
				     hssfCell5.setCellValue("设计完成月份   ");
				     
				     HSSFCell hssfCell6=hssfRow.createCell(6);
				     hssfCell6.setCellValue("专业审核人   ");
				     
				     HSSFCell hssfCell7=hssfRow.createCell(7);
				     hssfCell7.setCellValue("单项负责人   ");
				     
				     HSSFCell hssfCell8=hssfRow.createCell(8);
				     hssfCell8.setCellValue("概预算审核人   ");
				     
				     HSSFCell hssfCell9=hssfRow.createCell(9);
				     hssfCell9.setCellValue("概预算编制人   ");
				     
				     HSSFCell hssfCell10=hssfRow.createCell(10);
				     hssfCell10.setCellValue("详细站址   ");
				     
				     HSSFCell hssfCell11=hssfRow.createCell(11);
				     hssfCell11.setCellValue("覆盖区域名称   ");
				     
				     HSSFCell hssfCell12=hssfRow.createCell(12);
				     hssfCell12.setCellValue("经度 ");
				     
				     HSSFCell hssfCell13=hssfRow.createCell(13);
				     hssfCell13.setCellValue("纬度 ");
				     
				     HSSFCell hssfCell14=hssfRow.createCell(14);
				     hssfCell14.setCellValue("宏基站   ");
				     
				     HSSFCell hssfCell15=hssfRow.createCell(15);
				     hssfCell15.setCellValue("小基站   ");
				     
				     HSSFCell hssfCell16=hssfRow.createCell(16);
				     hssfCell16.setCellValue("拉远站   ");
				     
				     HSSFCell hssfCell17=hssfRow.createCell(17);
				     hssfCell17.setCellValue("信源站   ");
				     
				     HSSFCell hssfCell18=hssfRow.createCell(18);
				     hssfCell18.setCellValue("预立项文件  ");
				     
				     HSSFCell hssfCell19=hssfRow.createCell(19);
				     hssfCell19.setCellValue("BBU品牌  ");
				     
				     HSSFCell hssfCell20=hssfRow.createCell(20);
				     hssfCell20.setCellValue("BBU型号  ");
				     
				     HSSFCell hssfCell21=hssfRow.createCell(21);
				     hssfCell21.setCellValue("RRU品牌  ");
				     
				     HSSFCell hssfCell22=hssfRow.createCell(22);
				     hssfCell22.setCellValue("RRU型号  ");
				     
				     /*HSSFCell hssfCell23=hssfRow.createCell(23);
				     hssfCell23.setCellValue("天线品牌 ");
				     
				     HSSFCell hssfCell24=hssfRow.createCell(24);
				     hssfCell24.setCellValue("天线型号 ");*/
				     
				     HSSFCell hssfCell23=hssfRow.createCell(23);
				     hssfCell23.setCellValue("抗震设防烈度   ");
				     
				     HSSFCell hssfCell24=hssfRow.createCell(24);
				     hssfCell24.setCellValue("主设备安装方式   ");
				     
				     HSSFCell hssfCell25=hssfRow.createCell(25);
				     hssfCell25.setCellValue("工程类型   ");
				  
				     HSSFCell hssfCell26=hssfRow.createCell(26);
				     hssfCell26.setCellValue("天线方位角   ");
				     
				     HSSFCell hssfCell27=hssfRow.createCell(27);
				     hssfCell27.setCellValue("天线挂高  ");
				     
				     HSSFCell hssfCell28=hssfRow.createCell(28);
				     hssfCell28.setCellValue("总下倾角   ");
				     
				     HSSFCell hssfCell29=hssfRow.createCell(29);
				     hssfCell29.setCellValue("天馈情况   ");
				     
				     HSSFCell hssfCell30=hssfRow.createCell(30);
				     hssfCell30.setCellValue("配置   ");
				     
				     HSSFCell hssfCell31=hssfRow.createCell(31);
				     hssfCell31.setCellValue("RRU数量      ");
				     
				     HSSFCell hssfCell32=hssfRow.createCell(32);
				     hssfCell32.setCellValue("现网覆盖状况以及存在问题      ");
				     
				     int i;
				     for(i=0;i<33;i++){
				    	 HSSFCell hf = hssfRow.getCell(i);
				    	 hf.setCellStyle(style);
				    	 hssfSheet.autoSizeColumn(i);
				     }
			     }
			     
			     
			     
			     try{
			     FileOutputStream fileOutputStream=new FileOutputStream(text.getText()+"/"+aText+"-"+bText+"-"+cText+"-"+dText+"-"+Math.round(Math.random()*1000000)+".xls");
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
		
		JRadioButton d4 = new JRadioButton("中兴");
		d4.addActionListener(dListener);
		d.add(d4);
		GridBagConstraints gbc_d4 = new GridBagConstraints();
		gbc_d4.fill = GridBagConstraints.VERTICAL;
		gbc_d4.insets = new Insets(0, 0, 5, 5);
		gbc_d4.gridx = 4;
		gbc_d4.gridy = 6;
		frame.getContentPane().add(d4, gbc_d4);
		
		JLabel label_25 = new JLabel("");
		GridBagConstraints gbc_label_25 = new GridBagConstraints();
		gbc_label_25.fill = GridBagConstraints.BOTH;
		gbc_label_25.insets = new Insets(0, 0, 5, 0);
		gbc_label_25.gridx = 5;
		gbc_label_25.gridy = 6;
		frame.getContentPane().add(label_25, gbc_label_25);
		
		JLabel label_26 = new JLabel("");
		GridBagConstraints gbc_label_26 = new GridBagConstraints();
		gbc_label_26.fill = GridBagConstraints.BOTH;
		gbc_label_26.insets = new Insets(0, 0, 5, 5);
		gbc_label_26.gridx = 0;
		gbc_label_26.gridy = 7;
		frame.getContentPane().add(label_26, gbc_label_26);
		
			JLabel lblNewLabel_3 = new JLabel("请选择保存地址");
			GridBagConstraints gbc_lblNewLabel_3 = new GridBagConstraints();
			gbc_lblNewLabel_3.anchor = GridBagConstraints.BELOW_BASELINE;
			gbc_lblNewLabel_3.fill = GridBagConstraints.HORIZONTAL;
			gbc_lblNewLabel_3.insets = new Insets(0, 0, 5, 5);
			gbc_lblNewLabel_3.gridx = 1;
			gbc_lblNewLabel_3.gridy = 7;
			frame.getContentPane().add(lblNewLabel_3, gbc_lblNewLabel_3);
		
		
		
		text.setText("C:\\Users\\admin\\Desktop");
		GridBagConstraints gbc_text = new GridBagConstraints();
		gbc_text.anchor = GridBagConstraints.BELOW_BASELINE;
		gbc_text.insets = new Insets(0, 0, 5, 5);
		gbc_text.gridx = 2;
		gbc_text.gridy = 7;
		frame.getContentPane().add(text, gbc_text);
		
		JButton jButton=new JButton();
		jButton.setBackground(Color.WHITE);
		jButton.setForeground(Color.DARK_GRAY);
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
		GridBagConstraints gbc_jButton = new GridBagConstraints();
		gbc_jButton.anchor = GridBagConstraints.BELOW_BASELINE;
		gbc_jButton.insets = new Insets(0, 0, 5, 5);
		gbc_jButton.gridx = 3;
		gbc_jButton.gridy = 7;
		frame.getContentPane().add(jButton, gbc_jButton);
		
		JLabel label_27 = new JLabel("");
		GridBagConstraints gbc_label_27 = new GridBagConstraints();
		gbc_label_27.fill = GridBagConstraints.BOTH;
		gbc_label_27.insets = new Insets(0, 0, 5, 5);
		gbc_label_27.gridx = 4;
		gbc_label_27.gridy = 7;
		frame.getContentPane().add(label_27, gbc_label_27);
		
		JLabel label_28 = new JLabel("");
		label_28.setHorizontalAlignment(SwingConstants.CENTER);
		GridBagConstraints gbc_label_28 = new GridBagConstraints();
		gbc_label_28.fill = GridBagConstraints.BOTH;
		gbc_label_28.insets = new Insets(0, 0, 5, 0);
		gbc_label_28.gridx = 5;
		gbc_label_28.gridy = 7;
		frame.getContentPane().add(label_28, gbc_label_28);
		
		JLabel label_29 = new JLabel("");
		GridBagConstraints gbc_label_29 = new GridBagConstraints();
		gbc_label_29.fill = GridBagConstraints.BOTH;
		gbc_label_29.insets = new Insets(0, 0, 0, 5);
		gbc_label_29.gridx = 0;
		gbc_label_29.gridy = 8;
		frame.getContentPane().add(label_29, gbc_label_29);
		
		JLabel label_30 = new JLabel("");
		GridBagConstraints gbc_label_30 = new GridBagConstraints();
		gbc_label_30.fill = GridBagConstraints.BOTH;
		gbc_label_30.insets = new Insets(0, 0, 0, 5);
		gbc_label_30.gridx = 1;
		gbc_label_30.gridy = 8;
		frame.getContentPane().add(label_30, gbc_label_30);
		
		JLabel label_31 = new JLabel("");
		GridBagConstraints gbc_label_31 = new GridBagConstraints();
		gbc_label_31.fill = GridBagConstraints.BOTH;
		gbc_label_31.insets = new Insets(0, 0, 0, 5);
		gbc_label_31.gridx = 2;
		gbc_label_31.gridy = 8;
		frame.getContentPane().add(label_31, gbc_label_31);
		
		JLabel label_32 = new JLabel("");
		GridBagConstraints gbc_label_32 = new GridBagConstraints();
		gbc_label_32.fill = GridBagConstraints.BOTH;
		gbc_label_32.insets = new Insets(0, 0, 0, 5);
		gbc_label_32.gridx = 3;
		gbc_label_32.gridy = 8;
		frame.getContentPane().add(label_32, gbc_label_32);
		
		JLabel label_33 = new JLabel("");
		GridBagConstraints gbc_label_33 = new GridBagConstraints();
		gbc_label_33.fill = GridBagConstraints.BOTH;
		gbc_label_33.insets = new Insets(0, 0, 0, 5);
		gbc_label_33.gridx = 4;
		gbc_label_33.gridy = 8;
		frame.getContentPane().add(label_33, gbc_label_33);
		GridBagConstraints gbc_btnNewButton = new GridBagConstraints();
		gbc_btnNewButton.anchor = GridBagConstraints.BELOW_BASELINE_LEADING;
		gbc_btnNewButton.gridx = 5;
		gbc_btnNewButton.gridy = 8;
		frame.getContentPane().add(btnNewButton, gbc_btnNewButton);
		

		
		
	
	}
}
