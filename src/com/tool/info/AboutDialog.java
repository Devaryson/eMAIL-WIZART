package com.tool.info;

import java.awt.Cursor;
import java.awt.Desktop;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

//import email.code.ExcelToVcard;

import javax.swing.JLabel;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Color;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import javax.swing.ImageIcon;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.Toolkit;

@SuppressWarnings("serial")
public class AboutDialog extends JDialog {

	private final JPanel contentPanel = new JPanel();
	String labeltext;
	JFrame mf;

	String aboutTitle;

	public AboutDialog(JFrame parent, boolean demo, String labeltext) {

		super(parent, true);
		this.mf = parent;
		this.labeltext = labeltext;
		aboutTitle = ToolDetails.messageboxtitle;

		if (demo == true) {

			String center = ToolDetails.messageboxtitle;

			setTitle(center);

		} else {

			String center = ToolDetails.messageboxtitle;

			setTitle(center);

		}

		setIconImage(Toolkit.getDefaultToolkit().getImage(AboutDialog.class.getResource("/128x128.png")));
		setResizable(false);
		setBounds(100, 100, 480, 342);
		getContentPane().setLayout(null);
		contentPanel.setBackground(new Color(255, 255, 255));
		contentPanel.setBounds(0, 0, 464, 306);
		contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		getContentPane().add(contentPanel);
		contentPanel.setLayout(null);

		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setIcon(new ImageIcon(AboutDialog.class.getResource("/about.png")));
		lblNewLabel.setBounds(0, 1, 214, 308);
		contentPanel.add(lblNewLabel);

		JLabel lblNewLabel_1 = new JLabel(aboutTitle);
		lblNewLabel_1.setFont(new Font("Segoe UI Semibold", Font.BOLD, 13));
		lblNewLabel_1.setBounds(229, 3, 246, 28);
		contentPanel.add(lblNewLabel_1);

		JPanel panel = new JPanel();
		panel.setBackground(new Color(255, 255, 255));
		panel.setBounds(224, 72, 214, 198);
		contentPanel.add(panel);
		panel.setLayout(null);

		JLabel lblNewLabel_2 = new JLabel("Edition:              Standard ");
		lblNewLabel_2.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblNewLabel_2.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		lblNewLabel_2.setBounds(10, 10, 176, 24);
		panel.add(lblNewLabel_2);

		JLabel lblVersionStandard = new JLabel("Version:              21.9");
		lblVersionStandard.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblVersionStandard.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		lblVersionStandard.setBounds(10, 38, 183, 24);
		panel.add(lblVersionStandard);

		JLabel lblLicencseStandard = new JLabel("Licensed To:      " + labeltext);
		lblLicencseStandard.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblLicencseStandard.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		lblLicencseStandard.setBounds(10, 68, 203, 21);
		panel.add(lblLicencseStandard);

		JLabel websitelink = new JLabel("");
		websitelink.setBounds(10, 124, 163, 20);
		panel.add(websitelink);
		websitelink.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		websitelink.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		websitelink.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
				try {
					Desktop.getDesktop().browse(new URI("https://www.arysontechnologies.com"));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		websitelink.setForeground(new Color(0, 0, 205));
		websitelink.setText("Home");
		websitelink.setCursor(new Cursor(Cursor.HAND_CURSOR));

		JLabel lblSupportInformation = new JLabel("Support Information");
		lblSupportInformation.setBounds(5, 94, 181, 22);
		panel.add(lblSupportInformation);
		lblSupportInformation.setFont(new Font("Segoe UI", Font.BOLD, 12));

		JLabel supportlink = new JLabel("");
		supportlink.setBounds(10, 151, 164, 16);
		panel.add(supportlink);
		supportlink.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		supportlink.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		supportlink.setForeground(new Color(0, 0, 205));
		supportlink.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				try {
					Desktop.getDesktop().browse(
							new URI("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html"));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		supportlink.setText("Live Chat");
		supportlink.setCursor(new Cursor(Cursor.HAND_CURSOR));
		supportlink.setCursor(new Cursor(Cursor.HAND_CURSOR));

		JLabel saleslink = new JLabel("contact@arysontechnologies.com");
		saleslink.setBounds(10, 176, 197, 15);
		panel.add(saleslink);
		saleslink.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				try {
					Desktop.getDesktop().mail(new URI("mailto:contact@arysontechnologies.com" + ""));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		saleslink.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
		saleslink.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		saleslink.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
		saleslink.setForeground(new Color(0, 0, 205));
		{
			JButton okButton = new JButton("");
			okButton.setContentAreaFilled(false);
			okButton.setBorderPainted(false);
			okButton.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent e) {
					okButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/ok_about-hvr.png")));
				}

				@Override
				public void mouseExited(MouseEvent e) {
					okButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/ok_about.png")));
				}
			});
			okButton.setFocusPainted(false);
			okButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/ok_about.png")));
			okButton.setBounds(349, 274, 75, 29);
			contentPanel.add(okButton);
			okButton.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					dispose();
				}
			});
			okButton.setActionCommand("OK");
			getRootPane().setDefaultButton(okButton);
		}

		JButton userlicButton = new JButton("");
		userlicButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					Desktop.getDesktop()
							.browse(new URI("https://www.arysontechnologies.com/pdf/Eula%20for%20Aryson.pdf"));
				} catch (URISyntaxException | IOException ex) {
					// It looks like there's a problem
				}
			}
		});
		userlicButton.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				userlicButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/view-user-license-hvr.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				userlicButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/view-user-license.png")));
			}
		});
		userlicButton.setIcon(new ImageIcon(AboutDialog.class.getResource("/view-user-license.png")));
		userlicButton.setFocusPainted(false);
		userlicButton.setBorderPainted(false);
		userlicButton.setContentAreaFilled(false);
		userlicButton.setBounds(224, 274, 115, 27);
		contentPanel.add(userlicButton);

		JLabel lblProductInformation = new JLabel("Product Information");
		lblProductInformation.setFont(new Font("Segoe UI", Font.BOLD, 12));
		lblProductInformation.setBounds(225, 52, 181, 22);
		contentPanel.add(lblProductInformation);

		JLabel lblNewLabel_4 = new JLabel("Copyright (C) Aryson Technologies");
		lblNewLabel_4.setFont(new Font("Segoe UI", Font.PLAIN, 12));
		lblNewLabel_4.setBounds(235, 32, 202, 22);
		contentPanel.add(lblNewLabel_4);
	}
}
