package com.tool.activation;

import java.awt.Toolkit;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.border.EmptyBorder;
import javax.swing.text.DefaultEditorKit;
import javax.swing.text.JTextComponent;
import javax.swing.text.TextAction;

import com.main.EmailWizardApplication;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;
import javax.swing.Action;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.util.Base64;
import java.net.*;
import java.awt.event.ActionEvent;
import java.awt.Font;
import java.awt.Color;
import javax.swing.border.TitledBorder;

public class OnlineActivation extends JFrame {

	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	static JTextField textField_order_id;
	static JTextField textField_activation_key;
	private String order_id = "";
	private String activationKey = "";
	static JButton btnOnlineActivation;
	static JButton btnOfflineActivation;
	static HttpURLConnection conn = null;
	static File licFile;
	static BufferedWriter bufferedWriter = null;
	static JFrame frame;
	static int versiontype;
	public JButton btnBack;
	private JLabel label_1;
	DialogTests l;
	private JPanel panel;

	Starting_Frame sm = null;

	public OnlineActivation(JFrame startingframe, File file, boolean activatelocation) {

		super();
		sm = new Starting_Frame();
		if (activatelocation) {
			frame = (EmailWizardApplication) startingframe;

		} else {
			frame = (Starting_Frame) startingframe;
			activatelocation = false;

		}
		licFile = file;
		l = new DialogTests(frame);
		l.setLocationRelativeTo(null);
		l.setVisible(false);

		setResizable(false);
		setIconImage(Toolkit.getDefaultToolkit().getImage(ActivationFrame.class.getResource("/128x128.png")));
		setTitle(Starting_Frame.messageboxtitle);
		setBounds(100, 100, 679, 227);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JPopupMenu menu = new JPopupMenu();
		Action cut = new DefaultEditorKit.CutAction();
		cut.putValue(Action.NAME, "Cut");
		cut.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control X"));
		menu.add(cut);

		Action copy = new DefaultEditorKit.CopyAction();
		copy.putValue(Action.NAME, "Copy");
		copy.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control C"));
		menu.add(copy);

		Action paste = new DefaultEditorKit.PasteAction();
		paste.putValue(Action.NAME, "Paste");
		paste.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control V"));
		menu.add(paste);

		Action selectAll = new SelectAll();
		menu.add(selectAll);

		label_1 = new JLabel("");
		label_1.setIcon(new ImageIcon(ActivationFrame.class.getResource("/sidebar.gif")));
		label_1.setBounds(0, 0, 188, 200);
		contentPane.add(label_1);

		JLabel label = new JLabel("");
		label.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activation-bottom.png")));

		label.setBounds(0, 202, 597, 23);
		contentPane.add(label);

		panel = new JPanel();
		panel.setBackground(Color.WHITE);
		panel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBounds(198, 4, 465, 183);
		contentPane.add(panel);
		panel.setLayout(null);

		JLabel lblOrderId = new JLabel("Order No");
		lblOrderId.setBounds(0, 54, 104, 15);
		panel.add(lblOrderId);
		lblOrderId.setFont(new Font("Tahoma", Font.BOLD, 12));

		textField_order_id = new JTextField();
		textField_order_id.setBounds(105, 45, 324, 27);
		panel.add(textField_order_id);
		textField_order_id.setColumns(10);
		textField_order_id.setComponentPopupMenu(menu);

		JLabel lblA = new JLabel("Activation Key\r\n");
		lblA.setBounds(0, 93, 104, 27);
		panel.add(lblA);
		lblA.setFont(new Font("Tahoma", Font.BOLD, 11));

		textField_activation_key = new JTextField();
		textField_activation_key.setBounds(105, 93, 324, 27);
		panel.add(textField_activation_key);
		textField_activation_key.setColumns(10);
		textField_activation_key.setComponentPopupMenu(menu);

		btnOnlineActivation = new JButton("");
		btnOnlineActivation.setBounds(151, 140, 142, 38);
		panel.add(btnOnlineActivation);
		btnOnlineActivation.setRolloverEnabled(false);
		btnOnlineActivation.setRequestFocusEnabled(false);
		btnOnlineActivation.setOpaque(false);
		btnOnlineActivation.setFocusable(false);
		btnOnlineActivation.setFocusTraversalKeysEnabled(false);
		btnOnlineActivation.setFocusPainted(false);
		btnOnlineActivation.setDefaultCapable(false);
		btnOnlineActivation.setContentAreaFilled(false);
		btnOnlineActivation.setBorderPainted(false);
		btnOnlineActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/online-act-btn.png")));

		btnBack = new JButton("");
		btnBack.setBounds(114, 139, 45, 38);
		panel.add(btnBack);
		btnBack.setRolloverEnabled(false);
		btnBack.setRequestFocusEnabled(false);
		btnBack.setOpaque(false);
		btnBack.setFocusable(false);
		btnBack.setFocusTraversalKeysEnabled(false);
		btnBack.setFocusPainted(false);
		btnBack.setDefaultCapable(false);
		btnBack.setContentAreaFilled(false);
		btnBack.setBorderPainted(false);
		btnBack.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/act-back-btn.png")));

		btnOfflineActivation = new JButton("");
		btnOfflineActivation.setBounds(301, 143, 164, 29);
		panel.add(btnOfflineActivation);
		//btnOfflineActivation.setVisible(false);
		btnOfflineActivation.setRolloverEnabled(false);
		btnOfflineActivation.setRequestFocusEnabled(false);
		btnOfflineActivation.setOpaque(false);
		btnOfflineActivation.setFocusable(false);
		btnOfflineActivation.setFocusTraversalKeysEnabled(false);
		btnOfflineActivation.setFocusPainted(false);
		btnOfflineActivation.setDefaultCapable(false);
		btnOfflineActivation.setContentAreaFilled(false);
		btnOfflineActivation.setBorderPainted(false);
		btnOfflineActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/offline-act-btn.png")));
		btnOfflineActivation.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnOfflineActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/offline-act-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnOfflineActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/offline-act-btn.png")));
			}
		});
		
		btnOfflineActivation.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				ActivationFrame mf = new ActivationFrame();

				mf.setLocationRelativeTo(null);
				mf.setVisible(true);
				setVisible(false);
			}
		});
		btnBack.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnBack.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/act-back-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnBack.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/act-back-btn.png")));
			}
		});

		btnBack.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				dispose();
				Starting_Frame frame = new Starting_Frame();
				frame.setLocationRelativeTo(null);
				frame.setResizable(false);
				frame.setVisible(true);

			}
		});

		btnOnlineActivation.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnOnlineActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/online-act-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnOnlineActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/online-act-btn.png")));
			}
		});

		btnOnlineActivation.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {

				if (!(textField_activation_key.getText().isEmpty() && textField_activation_key.getText().isEmpty())) {

					l.setVisible(true);
					btnOnlineActivation.setEnabled(false);
					btnOfflineActivation.setEnabled(false);
					btnBack.setEnabled(false);

					Thread t = new Thread(new Runnable() {
						@Override
						public void run() {

							order_id = textField_order_id.getText().trim();
							activationKey = textField_activation_key.getText().trim();
							getResponseFromUrl(order_id, activationKey);

						}
					});
					t.start();
				} else {
					JOptionPane.showMessageDialog(null, "Order No or Activation key can not Empty!",
							Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));
					l.setVisible(false);

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}

			}
		});
	}

	public static void writedatainfile(File licFile, String orderId, String activationId, String type) {

		try {
			String BasicBase64format = null;
			PrintWriter clear = new PrintWriter(licFile);
			clear.print("");
			// other operations
			clear.close();

			FileWriter writer = new FileWriter(licFile, true);
			bufferedWriter = new BufferedWriter(writer);
			BasicBase64format = Base64.getEncoder().encodeToString(orderId.getBytes());
			bufferedWriter.write(BasicBase64format);
			bufferedWriter.newLine();
			BasicBase64format = Base64.getEncoder().encodeToString(activationId.getBytes());
			bufferedWriter.write(BasicBase64format);

		} catch (Exception e) {

			e.printStackTrace();
		} finally {
			try {
				bufferedWriter.close();
			} catch (IOException e) {

				e.printStackTrace();
			}
		}

	}

	public void getResponseFromUrl(String order_id, String activationKey) {
		try {

			// URL url = new
			// URL("http://www.howtomergepstfiles.com/mohit/activation-tool/outlook-pst-recovery-activation.php?order_no=2039119685667813A&activation_key=AbCD123&type=2");

			URL url = new URL(
					Starting_Frame.ToolUri + "order_no=" + order_id + "&activation_key=" + activationKey + "&type=1");
			conn = (HttpURLConnection) url.openConnection();
			conn.setUseCaches(false);
			conn.setDefaultUseCaches(false);
			conn.setRequestMethod("GET");
			System.out.println(conn.getUseCaches());

			String responsemessage = conn.getResponseMessage();
			if (responsemessage.equalsIgnoreCase("ok")) {

				InputStream inputStream = conn.getInputStream();
				BufferedReader in = new BufferedReader(new InputStreamReader(inputStream));
				String inputLine;
				String response = "";
				while ((inputLine = in.readLine()) != null) {

					response = inputLine;
					System.out.println(inputLine);
				}

				int urlresponse = Integer.valueOf(response);
				writedatainfile(licFile, order_id, activationKey, response);

				if (urlresponse == 1) {

					SwingUtilities.invokeLater(new Runnable() {

						@Override
						public void run() {
							dispose();
							l.setVisible(false);

							frame.setVisible(false);
							Starting_Frame.messageboxtitle = Starting_Frame.messageboxtitle + "(Single License)";

							JOptionPane.showMessageDialog(frame,
									Starting_Frame.messageboxtitle + "  Activated Successfully",
									Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(ActivationFrame.class.getResource("/check.png")));

							EmailWizardApplication mf = new EmailWizardApplication(false, urlresponse);
							mf.setLocationRelativeTo(null);
							mf.setVisible(true);

						}
					});

				}
				if (urlresponse == 2) {

					SwingUtilities.invokeLater(new Runnable() {

						@Override
						public void run() {

							dispose();
							l.setVisible(false);

							frame.setVisible(false);

							Starting_Frame.messageboxtitle = Starting_Frame.messageboxtitle + "(Admin License)";

							JOptionPane.showMessageDialog(frame,
									Starting_Frame.messageboxtitle + "  Activated Successfully",
									Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(ActivationFrame.class.getResource("/check.png")));

							EmailWizardApplication mf = new EmailWizardApplication(false, urlresponse);
							mf.setLocationRelativeTo(null);
							mf.setVisible(true);

						}
					});
				}
				if (urlresponse == 3) {

					SwingUtilities.invokeLater(new Runnable() {

						@Override
						public void run() {

							dispose();
							l.setVisible(false);
							frame.setVisible(false);

							Starting_Frame.messageboxtitle = Starting_Frame.messageboxtitle + "(Technical License)";
							JOptionPane.showMessageDialog(frame,
									Starting_Frame.messageboxtitle + "  Activated Successfully",
									Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(ActivationFrame.class.getResource("/check.png")));

							EmailWizardApplication mf = new EmailWizardApplication(false, urlresponse);
							mf.setLocationRelativeTo(null);
							mf.setVisible(true);

						}
					});
				}
				if (urlresponse == 3) {

					SwingUtilities.invokeLater(new Runnable() {

						@Override
						public void run() {

							dispose();
							l.setVisible(false);
							frame.setVisible(false);

							Starting_Frame.messageboxtitle = Starting_Frame.messageboxtitle + "(Enterprise License)";
							JOptionPane.showMessageDialog(frame,
									Starting_Frame.messageboxtitle + "  Activated Successfully",
									Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(ActivationFrame.class.getResource("/check.png")));

							EmailWizardApplication mf = new EmailWizardApplication(false, urlresponse);
							mf.setLocationRelativeTo(null);
							mf.setVisible(true);

						}
					});
				}

				if (urlresponse == 10) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Order No. can not Empty!", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					writedatainfile(licFile, "", "", "");

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}
				if (urlresponse == 11) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Activation Key can not Empty!", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					writedatainfile(licFile, "", "", "");

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}

				if (urlresponse == 12) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Type can not Empty!", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					writedatainfile(licFile, "", "", "");
					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}

				if (urlresponse == 13) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Please Enter valid Order No. and Activation key",
							Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					writedatainfile(licFile, "", "", "");

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}

				if (urlresponse == 14) {

					l.setVisible(false);
					// l.dispose();

					JOptionPane.showMessageDialog(null,
							"You can not activate remaining License Please contact " + sm.Salesemailid,
							Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					// frame.dispose();
					// dispose();
					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);

					// System.exit(0);
				}

				if (urlresponse == 15) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Sorry You have consumed all License",
							Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);

				}
				if (urlresponse == 16) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Error in Activation", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}

				if (urlresponse == 17) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, " Software Deactivate Successfully.",
							Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					writedatainfile(licFile, "", "", "");
					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}
				if (urlresponse == 18) {

					// type=2
					l.setVisible(false);
					JOptionPane.showMessageDialog(null, "Software Not Deactivate Due to some Error",
							Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}
				if (urlresponse == 19) {

					// type=3

					l.setVisible(false);
					JOptionPane.showMessageDialog(null, "Software Currently Active", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}
				if (urlresponse == 20) {

					// type3
					l.setVisible(false);
					JOptionPane.showMessageDialog(null, "Software Deactive", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}
				if (urlresponse == 21) {

					l.setVisible(false);

					JOptionPane.showMessageDialog(null, "Please Request Valid Type", Starting_Frame.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(OnlineActivation.class.getResource("/information.png")));

					btnOnlineActivation.setEnabled(true);
					btnOfflineActivation.setEnabled(true);
					btnBack.setEnabled(true);
				}
			} else {

				l.setVisible(false);
				JOptionPane.showMessageDialog(null, "Unable To Fetch Response From the Server",
						Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(OnlineActivation.class.getResource("/information.png")));

				btnOnlineActivation.setEnabled(true);
				btnOfflineActivation.setEnabled(true);
				btnBack.setEnabled(true);
			}
		} catch (MalformedURLException e1) {

			e1.printStackTrace();
		} catch (IOException e1) {

			l.setVisible(false);
			JOptionPane.showMessageDialog(null, "Not ready.Please check your Connection",
					Starting_Frame.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(OnlineActivation.class.getResource("/information.png")));

			btnOnlineActivation.setEnabled(true);
			btnOfflineActivation.setEnabled(true);
			btnBack.setEnabled(true);

		}
	}

	static class SelectAll extends TextAction {
		/**
		 * 
		 */
		private static final long serialVersionUID = 1L;

		public SelectAll() {
			super("Select All");
			putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control S"));
		}

		public void actionPerformed(ActionEvent e) {
			JTextComponent component = getFocusedComponent();
			component.selectAll();
			component.requestFocusInWindow();
		}
	}
}
