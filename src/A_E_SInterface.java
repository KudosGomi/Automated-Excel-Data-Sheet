package AESProgram;

import java.awt.EventQueue;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import javax.swing.JFrame;
import java.awt.BorderLayout;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.JTabbedPane;

public class A_E_SInterface {
	
	public JFrame theFrame; // mainFrame
	public mainPanel mainPanel;
	public JTabbedPane tab;
	
	public boolean changeTabs;
	private boolean pressedOnce = false;
	
	/* Start application */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					A_E_SInterface window = new A_E_SInterface();
					window.theFrame.setVisible(true);
				} catch (Exception e) {
					System.out.println("Main frame did not load.");
					e.printStackTrace();
				}
			}
		});
	}

	public A_E_SInterface() {
		initialize();
	}

	private void initialize() {
		
		theFrame = new JFrame();
		theFrame.setTitle("Automated Excel Sheet Program");
		theFrame.setBounds(100, 100, 700, 565);
		theFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		theFrame.setLayout(new BorderLayout());
		theFrame.add(createTabbedPane(), BorderLayout.CENTER);
		theFrame.setResizable(false);
		
		/* When changing tabs, the size resets or stays the same size on the main panel (first tab). */
		tab.addChangeListener(new ChangeListener() {
			public void stateChanged(ChangeEvent e) {
				
				if(tab.getSelectedIndex() == 1)
				{
					theFrame.setSize(711, 577);
				}
				else if(tab.getSelectedIndex() == 0)
				{
					if(mainPanel.pressedOnce == true)
					{
						theFrame.setSize(700, changeNamesPanel.mainFrameHeight); 
					}
				}
			}
		});
	}
	
	public JTabbedPane createTabbedPane() {
		
		tab = new JTabbedPane();
		mainPanel = new mainPanel(theFrame);
		createNewExcelPanel createNewExcelPanel = new createNewExcelPanel(A_E_SInterface.this);
		tab.addTab("Update CCR Excel File", null, mainPanel, null);
		tab.addTab("Create New CCR Excel File", null, createNewExcelPanel, null);
		return tab;
		
	}
}

