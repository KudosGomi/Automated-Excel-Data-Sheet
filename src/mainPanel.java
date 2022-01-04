package AESProgram;

import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.UIManager;

public class mainPanel extends JPanel {
	
	private ExcelOperations excelOperations = new ExcelOperations();
	public JFrame mainFrame;
	private JTextField textAvaya;
	private JTextField textCCRFile;
	private JTextField textSheetName;
	private JTextField textWeek;
	private JTextField textRowInsertData;
	private JTextField textSaveFileAs;
	private String filePathName1;
	private String filePathName2;
	private String filePathName3;
	public changeNamesPanel changeNamesPanel;
	public JButton generateNewExcelBtn;
	public JButton changeNameBtn;
	public boolean pressedOnce = false;
	
	public mainPanel(JFrame frame) {
		
		this.mainFrame = frame;
		setVisible(true);
		setBackground(new Color(255, 124, 0));
		setLayout(null);
		
		JLabel lblNewLabel = new JLabel("Enter Avaya file name:");
		lblNewLabel.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblNewLabel.setBounds(10, 123, 144, 32);
		add(lblNewLabel);
		
		JLabel lblSelectExcelFile = new JLabel("Select CCR Excel File:");
		lblSelectExcelFile.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblSelectExcelFile.setBounds(10, 80, 144, 32);
		add(lblSelectExcelFile);
		
		JLabel lblEnterSheetName = new JLabel("Enter sheet name:");
		lblEnterSheetName.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblEnterSheetName.setBounds(10, 209, 115, 32);
		add(lblEnterSheetName);
		
		JLabel lblEnterWeek = new JLabel("Enter week:");
		lblEnterWeek.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblEnterWeek.setBounds(10, 252, 86, 32);
		add(lblEnterWeek);
		
		JLabel lblEnterRowNumber = new JLabel("Enter row number to insert data:");
		lblEnterRowNumber.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblEnterRowNumber.setBounds(10, 331, 200, 32);
		add(lblEnterRowNumber);
		
		JLabel lblSaveFileAs = new JLabel("Save file as:");
		lblSaveFileAs.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblSaveFileAs.setBounds(10, 374, 86, 32);
		add(lblSaveFileAs);
		
		textCCRFile = new JTextField();
		textCCRFile.setBounds(164, 87, 397, 20);
		add(textCCRFile);
		textCCRFile.setColumns(10);
		
		textAvaya = new JTextField();
		textAvaya.setBounds(164, 132, 397, 20);
		add(textAvaya);
		textAvaya.setColumns(10);
		
		textSheetName = new JTextField();
		textSheetName.setBounds(135, 216, 426, 20);
		add(textSheetName);
		textSheetName.setColumns(10);
		
		textWeek = new JTextField();
		textWeek.setBounds(106, 260, 455, 20);
		add(textWeek);
		textWeek.setColumns(10);
		
		textRowInsertData = new JTextField();
		textRowInsertData.setBounds(220, 338, 341, 20);
		add(textRowInsertData);
		textRowInsertData.setColumns(10);
		
		textSaveFileAs = new JTextField();
		textSaveFileAs.setBounds(106, 381, 455, 20);
		add(textSaveFileAs);
		textSaveFileAs.setColumns(10);
		
		JButton btnBrowse1 = new JButton("Browse..");
		btnBrowse1.setBounds(585, 86, 89, 23);
		add(btnBrowse1);
		
		JButton btnBrowse2 = new JButton("Browse..");
		btnBrowse2.setBounds(585, 129, 89, 23);
		add(btnBrowse2);
		
		JButton btnBrowse3 = new JButton("Browse..");
		btnBrowse3.setBounds(585, 380, 89, 23);
		add(btnBrowse3);
		
		generateNewExcelBtn = new JButton("Generate CCR Excel File");
		generateNewExcelBtn.setBackground(UIManager.getColor("Button.foreground"));
		generateNewExcelBtn.setFont(new Font("Microsoft JhengHei UI", Font.BOLD, 12));
		generateNewExcelBtn.setBounds(249, 468, 185, 32);
		add(generateNewExcelBtn);
		
		JLabel lblMakeSureEach = new JLabel("(Make sure each person username or name is spelled correctly.)");
		lblMakeSureEach.setForeground(Color.white);
		lblMakeSureEach.setBounds(10, 49, 521, 20);
		add(lblMakeSureEach);
		
		JLabel lblenterRowsBetween = new JLabel("(Enter rows between 3 - 7 only.)");
		lblenterRowsBetween.setForeground(Color.white);
		lblenterRowsBetween.setBounds(10, 295, 521, 25);
		add(lblenterRowsBetween);
		
		JLabel lblMakeSureTo = new JLabel("Make sure to EXIT/CLOSE CCR EXCEL WINDOW when editing/updating.");
		lblMakeSureTo.setForeground(Color.BLUE);
		lblMakeSureTo.setFont(new Font("Tahoma", Font.BOLD | Font.ITALIC, 12));
		lblMakeSureTo.setHorizontalAlignment(SwingConstants.CENTER);
		lblMakeSureTo.setBounds(140, 12, 447, 14);
		add(lblMakeSureTo);
		
		JLabel lblpleaseMakeSure = new JLabel("(Please make sure the sheet name is correctly spelled.)");
		lblpleaseMakeSure.setForeground(Color.white);
		lblpleaseMakeSure.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblpleaseMakeSure.setBounds(10, 166, 521, 32);
		add(lblpleaseMakeSure);
		
		JButton helpBtn = new JButton("HELP");
		helpBtn.setBounds(598, 40, 64, 15);
		add(helpBtn);
		
		helpBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(null, "Make sure the usernames are correctly spelled in the Avaya file.\n     "
						+ "Make sure the usernames match the Avaya file in CCR file.\n"
						+ "                       REMINDER: Zero's affect total average!");
			}
		});
		
		btnBrowse1.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				
				JFileChooser fileChooser = new JFileChooser();
				if(fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
				{
					File file = fileChooser.getSelectedFile();
					filePathName1 = file.getAbsolutePath();
					String fileName = filePathName1.substring(filePathName1.lastIndexOf(File.separator) + 1, filePathName1.length());
					textCCRFile.setText(fileName);
				}
				
			}
		});
		
		btnBrowse2.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				
				JFileChooser fileChooser = new JFileChooser();
				if(fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
				{
					File file = fileChooser.getSelectedFile();
					filePathName2 = file.getAbsolutePath();
					String fileName = filePathName2.substring(filePathName2.lastIndexOf(File.separator) + 1, filePathName2.length());
					textAvaya.setText(fileName);	
				}	
				
			}
		});
		
		btnBrowse3.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setDialogTitle("Specify a file to save");
				int userSelection = fileChooser.showSaveDialog(null);
				if (userSelection == JFileChooser.APPROVE_OPTION) 
				{
				    File fileToSave = fileChooser.getSelectedFile();
				    filePathName3 = fileToSave.getAbsolutePath();
				    String fileName = filePathName3.substring(filePathName3.lastIndexOf(File.separator) + 1, filePathName3.length());
				    textSaveFileAs.setText(fileToSave.getAbsolutePath());
				}
				
			}
		});
		
		changeNameBtn = new JButton("Change username(s)");
		changeNameBtn.setFont(new Font("Lucida Sans Typewriter", Font.PLAIN, 11));
		changeNameBtn.setForeground(new Color(51, 0, 255));
		changeNameBtn.setBackground(UIManager.getColor("Button.foreground"));
		changeNameBtn.setBounds(263, 425, 160, 24);
		add(changeNameBtn);
		//mainFrame.getRootPane().setDefaultButton(generateNewExcelBtn);
		
		changeNameBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				if(pressedOnce == false)
				{
					generateNewExcelBtn.setLocation(254, 570);
					changeNamesPanel = new changeNamesPanel(mainFrame, generateNewExcelBtn, mainPanel.this);
					changeNamesPanel.setBounds(110, 470, 475, 80);
					add(changeNamesPanel);
					pressedOnce = true;
				}
				
			}
		});
		
		generateNewExcelBtn.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				
				int rowText = Integer.parseInt(textRowInsertData.getText());
				if(textCCRFile.getText().isEmpty()){
					JOptionPane.showMessageDialog(null, "Please enter the CCR file.");
				}
				else if(textAvaya.getText().isEmpty()){
					JOptionPane.showMessageDialog(null, "Please enter Avaya file.");
				}
				else if(textSheetName.getText().isEmpty()){
					JOptionPane.showMessageDialog(null, "Please enter sheet name.");
				}
				else if(textWeek.getText().isEmpty()){
					JOptionPane.showMessageDialog(null, "Please enter week.");
				}
				else if(textRowInsertData.getText().isEmpty()){
					JOptionPane.showMessageDialog(null, "Please enter row to insert data.");
				}
				else if(rowText > 7 || rowText < 3)
				{
					JOptionPane.showMessageDialog(null, "Enter rows only between 3 through 7.");
				}
				else if(textSaveFileAs.getText().isEmpty()){
					JOptionPane.showMessageDialog(null, "Please choose where to save file.");
				}
				else
				{
					if(pressedOnce == false)
					{
						excelOperations.getNamesFromCCRFile(filePathName1);
						excelOperations.readAvayaFile(filePathName2);
						excelOperations.createExcelFile(textSheetName.getText(), textRowInsertData.getText(), textWeek.getText(), textSaveFileAs.getText(), filePathName1);
					}
					else
					{
						excelOperations.getNamesFromCCRFile(filePathName1);
						excelOperations.readAvayaFile(filePathName2);
						excelOperations.nameChanges(changeNamesPanel.rowsList);
						excelOperations.createExcelFile(textSheetName.getText(), textRowInsertData.getText(), textWeek.getText(), textSaveFileAs.getText(), filePathName1);
					}
				}
			}
		});
	}
}