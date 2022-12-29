package com.tool.activation;

import java.awt.BorderLayout;
import java.awt.Color;
import javax.swing.ImageIcon;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.border.EtchedBorder;

import java.awt.SystemColor;
import java.awt.Font;

public class DialogTests extends JDialog {
    private static final long serialVersionUID = 1L;

//    public static void main(String[] args) {
//        try {
//            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        EventQueue.invokeLater(() -> {
//            new DialogTests(null).setVisible(true);
//        });
//    }

    public DialogTests(JFrame startingframe) {
        this.setAutoRequestFocus(false);
        this.setUndecorated(true);
        this.setAlwaysOnTop(true);
        this.setFocusableWindowState(true);
        this.setBackground(new Color(0,255,255,0));
        this.getContentPane().setLayout(new BorderLayout());
	
		
     

        JPanel contentPane = new JPanel();
        contentPane.setBackground(Color.WHITE);
    	contentPane.setBorder(new EtchedBorder(EtchedBorder.RAISED, SystemColor.textHighlight, null));
        setContentPane(contentPane);
        setBounds(200, 200, 305, 135);

        ImageIcon icon= new ImageIcon(DialogTests.class.getResource("/act.gif"));
        contentPane.setLayout(null);
        JLabel lblLoading = new JLabel(icon);
        lblLoading.setBounds(54, 2, 151, 122);
        lblLoading.setBackground(new Color(0, 0, 0, 0));
       getContentPane().add(lblLoading);
       
       JLabel lblNewLabel = new JLabel("Fetching data...");
       lblNewLabel.setForeground(Color.BLUE);
       lblNewLabel.setFont(new Font("Tahoma", Font.ITALIC, 11));
       lblNewLabel.setBounds(214, 117, 81, 14);
       contentPane.add(lblNewLabel);
    }

}