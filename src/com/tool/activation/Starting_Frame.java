package com.tool.activation;

import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.Toolkit;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.WindowConstants;
import javax.swing.border.EmptyBorder;



import net.Hash;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.sql.DriverManager;
import java.sql.Statement;
import java.util.Base64;
import java.awt.event.ActionEvent;
import javax.swing.JTextPane;
import java.awt.Font;
import java.awt.CardLayout;
import java.awt.Color;
import javax.swing.JLabel;
import javax.swing.border.TitledBorder;

import com.main.EmailWizardApplication;
import com.tool.info.ToolDetails;

public class Starting_Frame extends JFrame {

	public static boolean check=false;
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	static Starting_Frame frame;
	static int versiontype = 3;
	static File licFileon;
	static File licFileonoff;
	static String strSerialNumber;
	static String hashKey;
	static String defaultSerialNumber = "";
	String orderTFvalue = "";
	String infofilename = "";
	String productversion = "";
	String Salesemailid = "";

	public static String ToolUri = ToolDetails.ToolUri;
	public static String messageboxtitle = ToolDetails.messageboxtitle;
	public static String helpuri = ToolDetails.helpuri;
	public static String buyurl = ToolDetails.buyurl;

	static HttpURLConnection conn = null;
	public static String activationKey = "";
	public static String orderId = "";
	static String type = "";
	boolean activatefromedemo;
	public static EmailWizardApplication mf = null;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {

				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
					frame = new Starting_Frame();
				} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
						| UnsupportedLookAndFeelException e1) {

					e1.printStackTrace();
				}

				mf = new EmailWizardApplication(true, 3);


				File folder = null;
				if (System.getProperty("os.name").toLowerCase().contains("windows")) {
					folder = new File(System.getenv("APPDATA") + File.separator + messageboxtitle);
					folder.mkdirs();
				} else {
					folder = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
							+ "Application Support" + File.separator + messageboxtitle);
					folder.mkdirs();
				}

				String fileKey = null;
				if (System.getProperty("os.name").toLowerCase().contains("windows")) {
					licFileon = new File(System.getenv("APPDATA") + File.separator + messageboxtitle + File.separator
							+ "licenseOnline.lic");

					licFileonoff = new File(System.getenv("APPDATA") + File.separator + messageboxtitle + File.separator
							+ "license.lic");

				} else {
					licFileon = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
							+ "Application Support" + File.separator + messageboxtitle + File.separator
							+ "licenseOnline.lic");
					licFileonoff = new File(System.getenv("APPDATA") + File.separator + messageboxtitle + File.separator
							+ "license.lic");
				}

				if (licFileon.exists()) {
					try {
						FileReader fr = new FileReader(licFileon);
						BufferedReader br = new BufferedReader(new FileReader(licFileon));

						int count = 1;
						while ((fileKey = br.readLine()) != null) {
							System.out.println(fileKey);
							if (count == 1) {
								orderId = fileKey;
								byte[] actualByte = Base64.getDecoder().decode(orderId);
								orderId = new String(actualByte);

							} else if (count == 2) {
								activationKey = fileKey;

								byte[] actualByte = Base64.getDecoder().decode(activationKey);
								activationKey = new String(actualByte);

							}
							count++;

						}
						fr.close();
						br.close();

						if (!orderId.isEmpty() && !activationKey.isEmpty()) {

							SwingUtilities.invokeLater(new Runnable() {

								@Override
								public void run() {
									try {
										URL url = new URL(ToolUri + "order_no=" + orderId + "&activation_key="
												+ activationKey + "&type=3");
										conn = (HttpURLConnection) url.openConnection();
										conn.setUseCaches(false);
										conn.setDefaultUseCaches(false);
										conn.setRequestMethod("GET");
										System.out.println(conn.getUseCaches());

										String responsemessage = conn.getResponseMessage();
										int urlresponse = 0;
										if (responsemessage.equalsIgnoreCase("ok")) {

											InputStream inputStream = conn.getInputStream();
											BufferedReader in = new BufferedReader(new InputStreamReader(inputStream));
											String inputLine;
											String response = "";
											while ((inputLine = in.readLine()) != null) {

												response = inputLine;

											}

											String[] str = response.split(",");
											urlresponse = Integer.valueOf(str[0]);

											if (urlresponse == 19) {

												int version = Integer.valueOf(str[1]);
												mf = new EmailWizardApplication(false, version);
												mf.setLocationRelativeTo(null);
												mf.setVisible(true);

										
											} else if (urlresponse == 20) {

												JOptionPane.showMessageDialog(null, "Software Not Active",
														ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE, new ImageIcon(
																OnlineActivation.class.getResource("/cross-act.png")));

												frame.setLocationRelativeTo(null);
												frame.setResizable(false);
												frame.setVisible(true);

											}
										} else {

											JOptionPane.showMessageDialog(null,
													"Unable To Fetch Response From the Server", messageboxtitle,
													JOptionPane.INFORMATION_MESSAGE, new ImageIcon(
															OnlineActivation.class.getResource("/information.png")));

										}
									} catch (MalformedURLException e1) {

										e1.printStackTrace();
									} catch (IOException e1) {

										JOptionPane.showMessageDialog(null, "Not ready.Please check your Connection",
												messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
												new ImageIcon(OnlineActivation.class.getResource("/information.png")));

									}
								}
							});
						} else {

							try {
								frame.setLocationRelativeTo(null);
								frame.setResizable(false);
								frame.setVisible(true);
							} catch (Exception e) {
								e.printStackTrace();
							}

						}

					} catch (Exception ex) {

						ex.printStackTrace();
						try {
							frame.setLocationRelativeTo(null);
							frame.setResizable(false);
							frame.setVisible(true);
						} catch (Exception e) {
							e.printStackTrace();
						}

					}
				} else if (licFileonoff.exists()) {
					try {
						FileReader fr = new FileReader(licFileonoff);
						BufferedReader br = new BufferedReader(fr);
						fileKey = br.readLine();
						fr.close();
					} catch (Exception ex) {

					}

					if (fileKey != null) {

						strSerialNumber = ActivationFrame
								.getSerialNumber(System.getProperty("user.home").substring(0, 1));
						hashKey = new Hash().getHash(strSerialNumber);
						String licencetype=fileKey.substring(fileKey.length() - 1);
						fileKey=fileKey.substring(0, fileKey.length() - 1);
						
						 int intlic=Integer.valueOf(licencetype);
						 
						 
						 System.out.println(fileKey);
						
						if (hashKey.equals(fileKey)) {
							try {
								UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
							} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
									| UnsupportedLookAndFeelException e) {
								e.printStackTrace();
							}
							EmailWizardApplication mf = new EmailWizardApplication(false, intlic);
							mf.setLocationRelativeTo(null);
							mf.setVisible(true);
						} else {
							try {
								UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
							} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
									| UnsupportedLookAndFeelException e) {
								e.printStackTrace();
							}
							ActivationFrame af = new ActivationFrame();
							af.setLocationRelativeTo(null);
							af.setVisible(true);
						}
					} else {
						try {

							frame.setLocationRelativeTo(null);
							frame.setResizable(false);
							frame.setVisible(true);
						} catch (Exception e) {
							e.printStackTrace();
						}
					}

				} else {
					try {

						frame.setLocationRelativeTo(null);
						frame.setResizable(false);
						frame.setVisible(true);
					} catch (Exception e) {
						e.printStackTrace();
					}

				}

			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Starting_Frame() {
		setFont(new Font("Dialog", Font.BOLD, 11));

		setResizable(false);
		setIconImage(Toolkit.getDefaultToolkit().getImage(ActivationFrame.class.getResource("/128x128.png")));
		addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent arg0) {
				String warn = "Do you want to close?";
				int ans = JOptionPane.showConfirmDialog(frame, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.QUESTION_MESSAGE,
						new ImageIcon(ActivationFrame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					System.exit(0);
				}
			}
		});
		setTitle(messageboxtitle);
		setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
		setBounds(100, 100, 659, 207);
		contentPane = new JPanel();

		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		;

		JLabel label = new JLabel("");
		label.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activation-bottom.png")));

		label.setBounds(0, 156, 664, 23);
		contentPane.add(label);

		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBackground(Color.WHITE);
		panel.setBounds(521, 2, 132, 149);
		contentPane.add(panel);
		panel.setLayout(null);

		JButton btnNewButton = new JButton("");
		btnNewButton.setBounds(66, 5, 61, 68);
		panel.add(btnNewButton);

		btnNewButton.setRolloverEnabled(false);
		btnNewButton.setRequestFocusEnabled(false);
		btnNewButton.setOpaque(false);
		btnNewButton.setFocusable(false);
		btnNewButton.setFocusTraversalKeysEnabled(false);
		btnNewButton.setFocusPainted(false);
		btnNewButton.setDefaultCapable(false);
		btnNewButton.setContentAreaFilled(false);
		btnNewButton.setBorderPainted(false);
		btnNewButton.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/live-chat-btn.png")));

		JButton Buy_btn = new JButton("");
		Buy_btn.setBounds(0, 5, 68, 68);
		panel.add(Buy_btn);
		Buy_btn.setRolloverEnabled(false);
		Buy_btn.setRequestFocusEnabled(false);
		Buy_btn.setOpaque(false);
		Buy_btn.setFocusable(false);
		Buy_btn.setFocusTraversalKeysEnabled(false);
		Buy_btn.setFocusPainted(false);
		Buy_btn.setDefaultCapable(false);
		Buy_btn.setContentAreaFilled(false);
		Buy_btn.setBorderPainted(false);
		Buy_btn.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activation-buy-btn.png")));

		JButton btnActivate = new JButton("");
		btnActivate.setBounds(68, 78, 58, 61);
		panel.add(btnActivate);
		btnActivate.setRolloverEnabled(false);
		btnActivate.setRequestFocusEnabled(false);
		btnActivate.setOpaque(false);
		btnActivate.setFocusable(false);
		btnActivate.setFocusTraversalKeysEnabled(false);
		btnActivate.setFocusPainted(false);
		btnActivate.setDefaultCapable(false);
		btnActivate.setContentAreaFilled(false);
		btnActivate.setBorderPainted(false);
		btnActivate.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activate-btn.png")));

		JButton btnContinueDemo = new JButton("");
		btnContinueDemo.setBounds(0, 74, 68, 68);
		panel.add(btnContinueDemo);

		btnContinueDemo.setRolloverEnabled(false);
		btnContinueDemo.setRequestFocusEnabled(false);
		btnContinueDemo.setOpaque(false);
		btnContinueDemo.setFocusable(false);
		btnContinueDemo.setFocusTraversalKeysEnabled(false);
		btnContinueDemo.setFocusPainted(false);
		btnContinueDemo.setDefaultCapable(false);
		btnContinueDemo.setContentAreaFilled(false);
		btnContinueDemo.setBorderPainted(false);
		btnContinueDemo.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/demo-btn.png")));

		JPanel panel_1 = new JPanel();
		panel_1.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_1.setBackground(Color.WHITE);
		panel_1.setBounds(2, 2, 515, 149);
		contentPane.add(panel_1);
		panel_1.setLayout(null);

		JTextPane txtpane = new JTextPane();
		txtpane.setEditable(false);
		txtpane.setBounds(3, 6, 486, 48);

		// txtpane.setContentType("text/html");

		panel_1.add(txtpane);

		txtpane.setText(
				"Thanks for making  "+ToolDetails.messageboxtitle+" your first choice. As you are using the evaluation version with certain restrictions, we advise you to purchase the complete solution that offers full fledged solution.");
		txtpane.setBackground(Color.WHITE);
		txtpane.setFont(new Font("Arial", Font.BOLD, 12));

		JLabel label_1 = new JLabel("");
		label_1.setBounds(5, 59, 508, 79);
		label_1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/txt-act.png")));
		panel_1.add(label_1);

		btnContinueDemo.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnContinueDemo.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/demo-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnContinueDemo.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/demo-btn.png")));
			}
		});

		btnContinueDemo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				SwingUtilities.invokeLater(new Runnable() {

					@Override
					public void run() {
						File folder = null;
						if (System.getProperty("os.name").toLowerCase().contains("windows")) {
							folder = new File(System.getenv("APPDATA") + File.separator + ToolDetails.messageboxtitle);
							folder.mkdirs();
						} else {
							folder = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Application Support" + File.separator + ToolDetails.messageboxtitle);
							folder.mkdirs();
						}
						dispose();
						JOptionPane.showMessageDialog(null,
								ToolDetails.messageboxtitle + " will convert only 15 Items in trial version"
										+ System.lineSeparator()
										+ "Purchase licensed edition for converting all  data. ",
								ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
								new ImageIcon(EmailWizardApplication.class.getResource("/48x48.png")));
						mf = new EmailWizardApplication(true, ToolDetails.Lictype);
						mf.setLocationRelativeTo(null);
						mf.setVisible(true);
					

					}
				});

			}
		});
		btnActivate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnActivate.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activate-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnActivate.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activate-btn.png")));
			}
		});

		btnActivate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				dispose();
				OnlineActivation mf = new OnlineActivation(frame, licFileon, false);
				mf.setLocationRelativeTo(null);
				mf.setVisible(true);

				mf.addWindowListener(new WindowAdapter() {
					@Override
					public void windowClosing(WindowEvent arg0) {
						String warn = "Do you want to close?";
						int ans = JOptionPane.showConfirmDialog(frame, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
								JOptionPane.QUESTION_MESSAGE,
								new ImageIcon(EmailWizardApplication.class.getResource("/about-icon-2.png")));
						if (ans == JOptionPane.YES_OPTION) {
							setEnabled(true);
							// mf.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
							System.exit(0);
						} else {
							mf.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);

						}
					}
				});

			}
		});

		Buy_btn.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				Buy_btn.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activation-buy-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				Buy_btn.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/activation-buy-btn.png")));
			}
		});

		Buy_btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				openBrowser(buyurl);
			}
		});

		btnNewButton.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnNewButton.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/live-chat-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnNewButton.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/live-chat-btn.png")));
			}
		});

		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");
			}
		});
	}

	void openBrowser(String url) {
		if (Desktop.isDesktopSupported()) {
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.browse(new URI(url));
			} catch (IOException | URISyntaxException e) {
				// logger.warning("Warning : " + e.getMessage());
			}
		} else {
			Runtime runtime = Runtime.getRuntime();
			try {
				runtime.exec("xdg-open " + url);
			} catch (IOException e) {
				// logger.warning("Warning : " + e.getMessage());
			}
		}
	}
}
