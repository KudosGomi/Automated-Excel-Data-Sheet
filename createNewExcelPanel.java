package AESProgram;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.JScrollPane;

public class createNewExcelPanel extends JPanel {
	
	private ExcelOperations excelOperations = new ExcelOperations();
	private ExcelOperationsForNewBlankFile create = new ExcelOperationsForNewBlankFile();
	private JTextArea textArea;
	private JTextField sheetNameField;
	private JLabel savelabel = new JLabel("Save file as:");
	private JTextField saveCCRFile = new JTextField();
	private JButton browseSaveSpot = new JButton("Browse..");
	private Boolean checkedBlank = true;
	private JLabel enterCCRfile = new JLabel("Enter CCR file:");
	private JTextField getCCRFileField = new JTextField();
	private JButton browseCCRFile = new JButton("Browse..");
	private String filePathName;
	private String filePathName2;
	
	public createNewExcelPanel(A_E_SInterface tab) {

		setVisible(true);
		setBackground(Color.BLUE);
		setSize(690, 508);
		setLayout(null);
		
		JLabel msg = new JLabel("Enter names by separating them by a space or by entering each name on a new line.");
		msg.setForeground(Color.white);
		msg.setFont(new Font("Tahoma", Font.BOLD, 16));
		add(msg);
		
		JButton btnGenerateNewBlank = new JButton("Generate New Blank CCR Excel File");
		btnGenerateNewBlank.setBounds(0, 435, 690, 73);
		add(btnGenerateNewBlank);
		
		textArea = new JTextArea();
		textArea.setText("Enter new names here:");
		textArea.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e) {
				textArea.setText("");
			}
		});
		tab.tab.addChangeListener(new ChangeListener() {

			@Override
			public void stateChanged(ChangeEvent e) {
				
				if(tab.tab.getSelectedIndex() == 0)
				{
					textArea.setText("Enter new names here:");
				}
				
			}
		});
		textArea.setBounds(0, 45, 690, 388);
		add(textArea);
		
		JCheckBox chckbxNewBlankSheet = new JCheckBox("New Blank Sheet");
		chckbxNewBlankSheet.setForeground(new Color(255, 124, 0));
		chckbxNewBlankSheet.setBackground(Color.BLUE);
		chckbxNewBlankSheet.setBounds(0, 0, 140, 44);
		add(chckbxNewBlankSheet);
		
		JCheckBox chckbxNewBlankCcr = new JCheckBox("New Blank CCR Excel File");
		chckbxNewBlankCcr.setForeground(new Color(255, 124, 0));
		chckbxNewBlankCcr.setBackground(Color.BLUE);
		chckbxNewBlankCcr.setBounds(142, 0, 170, 44);
		add(chckbxNewBlankCcr);
		
		JLabel enterSheetName = new JLabel("Enter sheet name:");
		enterSheetName.setForeground(Color.WHITE);
		enterSheetName.setBounds(318, 15, 105, 14);
		enterSheetName.setEnabled(false);
		add(enterSheetName);
		
		sheetNameField = new JTextField();
		sheetNameField.setBounds(430, 14, 253, 20);
		sheetNameField.setColumns(10);
		sheetNameField.setBackground(Color.black);
		sheetNameField.setEnabled(false);
		add(sheetNameField);
		
		savelabel.setForeground(Color.WHITE);
		savelabel.setBounds(354, 63, 100, 14);
		add(savelabel);
		
		saveCCRFile.setBounds(433, 61, 165, 20);
		saveCCRFile.setColumns(10);
		saveCCRFile.setVisible(false);
		add(saveCCRFile);
		
		browseSaveSpot.setBounds(600, 61, 84, 20);
		add(browseSaveSpot);
		
		enterCCRfile.setForeground(Color.WHITE);
		enterCCRfile.setBounds(5, 63, 100, 14);
		enterCCRfile.setVisible(false);
		add(enterCCRfile);
		getCCRFileField.setBounds(95, 61, 160, 20);
		getCCRFileField.setHorizontalAlignment(SwingConstants.TRAILING);
		getCCRFileField.setColumns(10);
		getCCRFileField.setVisible(false);
		add(getCCRFileField);
		browseCCRFile.setBounds(257, 61, 84, 20);
		browseCCRFile.setVisible(false);
		add(browseCCRFile);
		
		chckbxNewBlankSheet.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {

				if(checkedBlank)
				{
					chckbxNewBlankCcr.setEnabled(false);
					enterSheetName.setEnabled(true);
					sheetNameField.setEnabled(true);
					sheetNameField.setBackground(Color.white);
					textArea.setBounds(0, 90, 690, 343);
					enterCCRfile.setVisible(true);
					getCCRFileField.setVisible(true);
					browseCCRFile.setVisible(true);
					saveCCRFile.setVisible(true);
					checkedBlank = false;
				}
				else 
				{
					chckbxNewBlankCcr.setEnabled(true);
					enterSheetName.setEnabled(false);
					sheetNameField.setEnabled(false);
					sheetNameField.setBackground(Color.black);
					textArea.setBounds(0, 45, 690, 388);
					checkedBlank = true;
				}
			}
		});
		
		chckbxNewBlankCcr.addActionListener(new ActionListener() {
			
			Boolean checked = true;
			
			public void actionPerformed(ActionEvent e) {
				
				if(checked)
				{
					chckbxNewBlankSheet.setEnabled(false);
					enterSheetName.setEnabled(true);
					sheetNameField.setEnabled(true);
					sheetNameField.setBackground(Color.white);
					textArea.setBounds(0, 90, 690, 343);
					saveCCRFile.setVisible(true);
					browseSaveSpot.setVisible(true);
					enterCCRfile.setVisible(false);
					getCCRFileField.setVisible(false);
					browseCCRFile.setVisible(false);
					checked = false;
				}
				else 
				{
					chckbxNewBlankSheet.setEnabled(true);
					enterSheetName.setEnabled(false);
					sheetNameField.setEnabled(false);
					sheetNameField.setBackground(Color.black);
					textArea.setBounds(0, 45, 690, 388);
					checked = true;
				}
				
			}
		});
		
		browseCCRFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				JFileChooser fileChooser = new JFileChooser();
				if(fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
				{
					File file = fileChooser.getSelectedFile();
					filePathName = file.getAbsolutePath();
					String fileName = filePathName.substring(filePathName.lastIndexOf(File.separator) + 1, filePathName.length());
					getCCRFileField.setText(fileName);	
				}
				
			}
		});
		
		browseSaveSpot.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setDialogTitle("Specify a file to save");
				int userSelection = fileChooser.showSaveDialog(null);
				if (userSelection == JFileChooser.APPROVE_OPTION) 
				{
				    File fileToSave = fileChooser.getSelectedFile();
				    filePathName2 = fileToSave.getAbsolutePath();
				    String fileName = filePathName2.substring(filePathName2.lastIndexOf(File.separator) + 1, filePathName2.length());
				    saveCCRFile.setText(fileToSave.getAbsolutePath());
				}
				
			}
		});
		
		btnGenerateNewBlank.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {

				create.getNamesFromTextArea(textArea);
				if(chckbxNewBlankSheet.isEnabled() && chckbxNewBlankCcr.isEnabled())
				{
					JOptionPane.showMessageDialog(null, "Please select a new blank sheet or a new blank file.");
				}
				else 
				{
					create.createNewBlankExcelWithNames(chckbxNewBlankSheet, chckbxNewBlankCcr, sheetNameField.getText(), filePathName, saveCCRFile);
				}

			}
		});
	}
}