package com.cm.oe.ui;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import com.cm.oe.budget.gen.BudgetWriter1;

public class MainPathCreate {

	private JFrame frame;
	public JTextField aText = null;
	public JTextField bText =null;
	public JTextField cText =null;
	public JTextField dText =null;

	/**
	 * Launch the application.
	 */

	/**
	 * Create the application.
	 */
	public MainPathCreate() {
		initialize();
		frame.setVisible(true);
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setResizable(false);
		frame.setBounds(100, 100, 660, 425);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		

		
	     aText = new JTextField();
		aText.setText("C:\\Users\\admin\\Desktop");
		aText.setBounds(220, 39, 145, 24);
		frame.getContentPane().add(aText);
		
		bText = new JTextField();
		bText.setText("C:\\Users\\admin\\Desktop");
		bText.setBounds(221, 102, 144, 24);
		frame.getContentPane().add(bText);
		
		 cText = new JTextField();
		cText.setText("C:\\Users\\admin\\Desktop");
		cText.setBounds(223, 160, 142, 24);
		frame.getContentPane().add(cText);
		
		 dText = new JTextField();
		dText.setText("C:\\Users\\admin\\Desktop\\");
		dText.setBounds(223, 220, 142, 24);
		frame.getContentPane().add(dText);
		
		JButton aButton = new JButton("...");
		aButton.setBounds(426, 39, 93, 23);
		aButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				 JFileChooser jFileChooser=new JFileChooser();
		            ExcelFileFilter ef = new ExcelFileFilter();
		            jFileChooser.addChoosableFileFilter(ef);
		            jFileChooser.setFileFilter(ef);
		            int state = jFileChooser.showOpenDialog(null); 
		            if (state == 1) {  
		                return;  
		            } else {  
		                File file = jFileChooser.getSelectedFile(); 
		                aText.setText(file.getAbsolutePath());  
		            }   
			}
		});
		frame.getContentPane().add(aButton);
		
		JButton bButton = new JButton("...");
		bButton.setBounds(426, 102, 93, 23);
		bButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				 JFileChooser jFileChooser=new JFileChooser();
				 ExcelFileFilter ef = new ExcelFileFilter();
		            jFileChooser.addChoosableFileFilter(ef);
		            jFileChooser.setFileFilter(ef);
				 int state = jFileChooser.showOpenDialog(null); 
		            if (state == 1) {  
		                return;  
		            } else {  
		                File file = jFileChooser.getSelectedFile(); 
		                bText.setText(file.getAbsolutePath());  
		            }   
			}
		});
		frame.getContentPane().add(bButton);
		
		JButton cButton = new JButton("...");
		cButton.setBounds(426, 160, 93, 23);
		cButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				 JFileChooser jFileChooser=new JFileChooser();
				 ExcelFileFilter ef = new ExcelFileFilter();
		            jFileChooser.addChoosableFileFilter(ef);
		            jFileChooser.setFileFilter(ef);
				 int state = jFileChooser.showOpenDialog(null); 
		            if (state == 1) {  
		                return;  
		            } else {  
		                File file = jFileChooser.getSelectedFile(); 
		                cText.setText(file.getAbsolutePath());  
		            }   
			}
		});
		frame.getContentPane().add(cButton);
		
		JButton dButton = new JButton("...");
		dButton.setBounds(426, 220, 93, 23);
		dButton.addActionListener(new ActionListener() {
			
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
		                dText.setText(file.getAbsolutePath());  
		            }   
			}
		});
		frame.getContentPane().add(dButton);
		
		JButton confirmButton = new JButton("确认路径选择");
		confirmButton.setBounds(255, 319, 178, 23);
	    confirmButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				if(aText.getText()==""||bText.getText()==""||cText.getText()==""||dText.getText()==""){
					JOptionPane.showMessageDialog(null, "请填写全部路径");
					return;
				}else{
					int state= JOptionPane.showConfirmDialog(null,"确定选择的路径？");
					if(state==0){
						boolean flag = false;
						String filePath = "testfiles/templete_allnew.docx";
						String path1 = bText.getText();
						String path2 = cText.getText();
						String output = dText.getText()+"/";
						String tablePath = "testfiles/tables.xls";
						String excelPath = aText.getText();
						try {
							BudgetWriter1 xwpf = new BudgetWriter1( path1,  path2,  filePath,  tablePath,  excelPath,
									 output);
							xwpf.testReadByDoc();
							flag = true;
							if(flag == true){
								JOptionPane.showMessageDialog(null, "生成完毕！");
								System.exit(0);
							}
						} catch (Exception e1) {
							e1.printStackTrace();
						}
					}else{
						return;
					}
				}
				
			}
		});
		frame.getContentPane().add(confirmButton);
		
		JLabel lblNewLabel = new JLabel("一体化基站勘察汇总表\r\n");
		lblNewLabel.setBounds(57, 43, 132, 15);
		frame.getContentPane().add(lblNewLabel);
		
		JLabel lblg = new JLabel("4G工程基站预算表路径");
		lblg.setBounds(57, 106, 132, 15);
		frame.getContentPane().add(lblg);
		
		JLabel lblgg = new JLabel("3G4G工程基站预算表路径");
		lblgg.setBounds(57, 164, 132, 15);
		frame.getContentPane().add(lblgg);
		
		JLabel label_2 = new JLabel("文件生成路径");
		label_2.setBounds(57, 224, 121, 15);
		frame.getContentPane().add(label_2);
		
		
		 
	}
}
