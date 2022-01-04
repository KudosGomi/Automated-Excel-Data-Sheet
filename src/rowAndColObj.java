package AESProgram;

import java.awt.Color;
import java.awt.Font;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingConstants;

public class rowAndColObj extends JPanel {
	
	private JTextField enterNameText;
	private JTextField enterColText;
	
	public rowAndColObj(int Yposition, JFrame mainFrame, int mainFrameHeight, JPanel addNamesPanel, int panelHeight, int numOfNames, JButton theAddBtn, int btnHeight, JButton genBtn, int genBtnHeight) {
		
		mainFrame.setSize(700, mainFrameHeight);
		addNamesPanel.setSize(475, panelHeight);
		genBtn.setLocation(254, genBtnHeight);
		theAddBtn.setLocation(178, btnHeight);
		
		JLabel enterNameLabel = new JLabel(numOfNames + ".)  Enter name:");
		enterNameLabel.setHorizontalAlignment(SwingConstants.CENTER);
		enterNameLabel.setBounds(40, Yposition, 88, 16);
		enterNameLabel.setFont(new Font("Times New Roman", Font.PLAIN, 13));
		enterNameLabel.setForeground(Color.WHITE);
		addNamesPanel.add(enterNameLabel);
		enterNameText = new JTextField();
		enterNameText.setBounds(136, Yposition, 100, 18);
		enterNameText.setColumns(10);
		addNamesPanel.add(enterNameText);
		JLabel enterColLabel = new JLabel("Enter column:");
		enterColLabel.setHorizontalAlignment(SwingConstants.CENTER);
		enterColLabel.setBounds(267, Yposition, 80, 16);
		enterColLabel.setFont(new Font("Times New Roman", Font.PLAIN, 13));
		enterColLabel.setForeground(Color.WHITE);
		addNamesPanel.add(enterColLabel);
		enterColText = new JTextField();
		enterColText.setBounds(352, Yposition, 55, 18);
		enterColText.setColumns(10);
		enterColText.setHorizontalAlignment(SwingConstants.TRAILING);
		addNamesPanel.add(enterColText);
		
	}
	
	public String getName() {
		return enterNameText.getText();
	}
	
	public String getCol() {
		return enterColText.getText();
	}

}
