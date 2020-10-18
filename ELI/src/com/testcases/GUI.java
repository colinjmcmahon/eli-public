package com.testcases;

import java.awt.FlowLayout;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.swing.JOptionPane;
import javax.swing.JLabel;

public class GUI extends JFrame{

	private JLabel Label1,Label2;
	
	public GUI(){
		super ("Easy Language Interface");

		//This bit here we need to work on - not integrated yet
		setLayout (new FlowLayout());
		Label1 = new JLabel("ELI Test Automation Framework");
		Label2 = new JLabel("Please select Datasheet to use");
		Label1.setToolTipText("Please select which data spreadsheet to use");
		add(Label1);
       //------------------------------------------------------
		
		String AskUser = JOptionPane.showInputDialog("Select Option 1 or 2");
        
        int response = Integer.parseInt(AskUser);
        
       
	}
}
