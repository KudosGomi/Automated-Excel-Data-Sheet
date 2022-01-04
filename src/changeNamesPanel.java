package AESProgram;

import java.awt.Color;


import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;

import org.apache.poi.ss.formula.atp.AnalysisToolPak;

public class changeNamesPanel  extends JPanel{
	
	public ArrayList<rowAndColObj> rowsList  = new ArrayList<rowAndColObj>();
	
	private JButton addAnotherNameBtn;
	public static int numOfNames = 0;
	private static int Yposition = 20;
	public static int mainFrameHeight = 673;
	private static int btnHeight = 52;
	private static int panelheight = 94;
	private static int genBtnHeight = 578;
	
	public changeNamesPanel(JFrame mainFrame, JButton generateNewExcelBtn, mainPanel mainPane) {
		
		setVisible(true);
		setBackground(Color.BLUE);
		setLayout(null);
		addAnotherNameBtn = new JButton("Add name");
		addAnotherNameBtn.setBounds(180, 47, 110, 20);
		add(addAnotherNameBtn);
		addNameAndColOBJ(mainFrame, generateNewExcelBtn, mainPane);
		
	}
	
	private void addNameAndColOBJ(JFrame mainFrame, JButton generateBtn, mainPanel mainPane) {
		
		rowAndColObj firstRow = new rowAndColObj(Yposition, mainFrame, mainFrameHeight, 
				changeNamesPanel.this, panelheight, ++numOfNames, 
				addAnotherNameBtn, btnHeight, generateBtn, genBtnHeight);
		rowsList.add(firstRow);
		add(firstRow);
		updateBounds(24, 24, 24, 24, 24);
		
		addAnotherNameBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				if(rowsList.size() < 10)
				{
					rowAndColObj addNewRow = new rowAndColObj(Yposition, mainFrame, mainFrameHeight, 
							changeNamesPanel.this, panelheight, ++numOfNames, 
							addAnotherNameBtn, btnHeight, generateBtn, genBtnHeight);
					rowsList.add(addNewRow);
					add(addNewRow);
					updateBounds(24, 24, 24, 24, 24);
				}
				
			}
		});
		
	}
	
	public void updateBounds(int Ypos, int mFHeight, int panH, int btnH, int genBtnH) {

		Yposition = Yposition + Ypos;
		mainFrameHeight = mainFrameHeight + mFHeight;
		panelheight = panelheight + panH;
		btnHeight = btnHeight + btnH;
		genBtnHeight = genBtnHeight + genBtnH;

	}

}
