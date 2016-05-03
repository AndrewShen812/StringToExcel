package com.gwcd.sy.excel;

import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JTextPane;

public class HomePage{

	private JFrame frame;
	private JTextField textOrgPath;
	private JTextField textDstPath;
	private JTextPane textInfo;
	private JButton btnBrowsePath;
	private JButton btnBrowseDst;
	private JButton btnWriteExcel;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					HomePage window = new HomePage();
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
	public HomePage() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 500, 360);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel lblNewLabel = new JLabel("\u8DEF\u5F84\uFF1A");
		lblNewLabel.setBounds(10, 10, 54, 15);
		frame.getContentPane().add(lblNewLabel);
		
		textOrgPath = new JTextField();
		textOrgPath.setBounds(53, 7, 351, 21);
		textOrgPath.setText("E:\\workspace\\Android\\QrCodeTools\\res\\values\\strings.xml");
		frame.getContentPane().add(textOrgPath);
		textOrgPath.setColumns(10);
		
		btnBrowsePath = new JButton("\u6D4F\u89C8");
		btnBrowsePath.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				showFileChooseDialog();
			}
		});
		btnBrowsePath.setBounds(414, 6, 60, 23);
		frame.getContentPane().add(btnBrowsePath);
		
		textInfo = new JTextPane();
		textInfo.setBounds(10, 35, 464, 200);
		frame.getContentPane().add(textInfo);
		
		btnWriteExcel = new JButton("\u751F\u6210excel\u6587\u4EF6");
		btnWriteExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				WriteToExcel();
			}
		});
		btnWriteExcel.setBounds(363, 277, 111, 23);
		frame.getContentPane().add(btnWriteExcel);
		
		JLabel label = new JLabel("\u6587\u4EF6\u8DEF\u5F84\uFF1A");
		label.setBounds(10, 249, 60, 15);
		frame.getContentPane().add(label);
		
		textDstPath = new JTextField();
		textDstPath.setBounds(80, 246, 324, 21);
		textDstPath.setText("E:\\e.xls");
		frame.getContentPane().add(textDstPath);
		textDstPath.setColumns(10);
		
		btnBrowseDst = new JButton("\u6D4F\u89C8");
		btnBrowseDst.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				showDstPathDialog();
			}
		});
		btnBrowseDst.setBounds(414, 245, 60, 23);
		frame.getContentPane().add(btnBrowseDst);
	}
	
	private void showFileChooseDialog() {
		JFileChooser jfc=new JFileChooser();  
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );  
        jfc.showDialog(new JLabel(), "选择");  
        File file=jfc.getSelectedFile();
        if (file != null) {
        	if(file.isDirectory()){  
        		textInfo.setText("请选择有效的xml资源文件");
                System.out.println("文件夹:"+file.getAbsolutePath());  
            }else if(file.isFile()){
            	String name = file.getName();
            	if (name.endsWith(".xml")) {
            		textInfo.setText("文件有效");
            		textOrgPath.setText(file.getAbsolutePath());
            	} else {
            		textInfo.setText("请选择有效的xml资源文件");
            	}
            }  
        }
	}
	
	private void showDstPathDialog() {
		JFileChooser jfc=new JFileChooser();  
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );  
        jfc.showDialog(new JLabel(), "选择");  
        File file=jfc.getSelectedFile();
        if (file != null) {
        	if(file.isDirectory()){  
        		textInfo.setText("请选择有效的Excel文件");
                System.out.println("文件夹:"+file.getAbsolutePath());  
            }else if(file.isFile()){
            	String name = file.getName();
            	if (name.endsWith(".xls")) {
            		textInfo.setText("目标文件有效");
            		textOrgPath.setText(file.getAbsolutePath());
            	} else {
            		textInfo.setText("请选择有效的Excel资源文件");
            	}
            }  
        }
	}
	
	private void WriteToExcel() {
		String org = textOrgPath.getText();
		String dst = textDstPath.getText();
		if (org == null || org.isEmpty()) {
			textInfo.setText("请选择resource文件！");
			return;
		}
		if (dst != null && !dst.isEmpty() && dst.endsWith(".xls")) {
			ExcelUtils parser = ExcelUtils.getInstance();
			parser.doWiriteExcel(org, dst);
		} else {
			textInfo.setText("目标文件无效！");
		}
	}
}
