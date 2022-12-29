package com.main;

import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Point;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.font.TextAttribute;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.security.GeneralSecurityException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TimeZone;

import javax.swing.Action;
import javax.swing.ButtonGroup;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFormattedTextField;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JPopupMenu;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.KeyStroke;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.MatteBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;
import javax.swing.text.DefaultEditorKit;
import javax.swing.text.JTextComponent;
import javax.swing.text.NumberFormatter;
import javax.swing.text.TextAction;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;

import org.slf4j.LoggerFactory;

import com.api.ews.EWSOffice;
import com.api.google.GoogleLogin;
import com.aspose.cells.Cell;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.email.EmailClient;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapException;
import com.aspose.email.ImapFolderInfo;
import com.aspose.email.ImapFolderInfoCollection;
import com.aspose.email.ImapNamespace;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.MultiConnectionMode;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SecurityOptions;
import com.constants.InputSource;
import com.constants.OutputSource;
import com.download.email.GmailBackup;
import com.download.email.ImapEmailBackUp;
import com.download.email.MsOfficeBackup;
import com.downoad.googleapp.CalenderBackup;
import com.downoad.googleapp.ContactBackup;
import com.downoad.googleapp.DriveBackup;
import com.downoad.googleapp.PhotosBackup;
import com.exceptions.ExceptionHandler;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.Label;
import com.google.api.services.gmail.model.ListLabelsResponse;
import com.toedter.calendar.JDateChooser;
import com.toedter.calendar.JTextFieldDateEditor;
import com.tool.activation.AsposeActivation;
import com.tool.activation.OnlineActivation;
import com.tool.info.AboutDialog;
import com.tool.info.ToolDetails;
import com.util.CSVUtils;
import com.util.FileNamingUtils;
import com.util.LogUtils;

import it.cnr.imaa.essi.lablib.gui.checkboxtree.CheckboxTree;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.remote.ServiceRequestException;
import microsoft.exchange.webservices.data.core.exception.service.remote.ServiceResponseException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.search.FolderView;
import javax.swing.JTextArea;

public class EmailWizardApplication extends JFrame {
	Credential inputGmailCredential;
	Credential outputGmailCredential;
	Gmail outputGmailService;
	private String user = "me";
	final static String SELCTION_TYPE_INPUT = "input";
	final static String SELCTION_TYPE_OUTPUT = "output";
	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	
	Label parentLabel;
	List<Label> lableList;
	public static JCheckBox c_DeepFolderTraversal;
	static EWSOffice ews;
	GoogleLogin googleLogin;
	static ExchangeService service;
	JLabel lblServiceaccountIdAndHostName;
	JLabel lblPassworduser;
	JLabel lblPFileAndPortNumber;
	JLabel lblUsername_1;
	public static String selectedInput = InputSource.GMAIL.getValue();
	static String inputIMAPHostName = "imap.gmail.com";
	static int inputIMAPPortNo = 993;
	DefaultTableModel model;
	int count = 0;

	public static ImapClient clientforimap_input;
	public static IConnection iconnforimap_input;

	public static ImapClient clientforimap_Output;
	public static IConnection iconnforimap_Output;

	String imapClientFolderpath;
	boolean isGmailDefaultFolderCreated = false;

	public static String output_password;
	public static String output_userName;
	public static String output_imapHost;
	public static String output_portNo;

	public static String input_password;
	public static String input_userName;

	public static String imapFolderPath;
	public static IEWSClient office365Client;
	public static long checkImapConnectionTime;

	public static String portNo;
	JPanel SavingOptionPanel;
	JLabel lblimapGif;

	public static JRadioButton radioButtonMB;
	public static JRadioButton rdbtnGb;
	public static JCheckBox chckbxSkipDuplicate;
	JPanel panel_5;
	JCheckBox c_contact;
	JCheckBox c_drive;
	JCheckBox c_calendar;
	public static JRadioButton rdbtnDateFilter;
	private static final long serialVersionUID = 1L;
	JCheckBox c_photos;
	public static JPanel CardLayout;
	private JPanel contentPane;
	private JTextField InputtxtUserName;
	private JTextField txtServiceAccountIDorImapHostName;
	private static JTextField textField_p12FileAndPortNo;
	public static JTable table_UserDetails;
	public static JTextField textField_DownloadingPath;
	public static String detinationPath;
	public static JTable table_Downloading;
	public static DefaultTableModel modelDownloading;
	public static int rownCount;
	JButton btnBack_p2;
	JCheckBox c_email;
	JButton btnBrowseP2File;
	JButton btn_Login;
	JButton btnNext_p2;
	JButton btnBack_p3;
	JButton btnNew_p3;
	JButton btnDownloadingPath;
	JButton btnDownloading;
	JButton btnBack_p4;
	private JLabel l_topbar;
	private JLabel l_drive;
	private JLabel l_contact;
	private JLabel l_calendar;
	private JLabel l_photos;
	private JLabel l_gmail;

	private JLabel label_LoginGif;
	private JPanel LoginPanel_1;
	private JLabel lblUserName;
	public static JProgressBar progressBar_Downloading;
	public static JLabel lblNoInternetConnection;
	public static JLabel lblDownloading;
	public static boolean stop;

	public static boolean demo = true;
	JButton btnStop;
	private JLabel label_8;
	private JLabel label_9;
	public static JLabel downloadingFileName;
	private JLabel dataName;
	private JPanel ProgressBarPanel;
	private final ButtonGroup buttonGroup = new ButtonGroup();
	static JPanel SavingOptionPanel_1;
	private JLabel l_bottombar;
	public static JDateChooser start_dateChooser;
	public static JDateChooser end_dateChooser;
	private JPanel p_outputSavingformat;
	JButton btn_login;

	JLabel lblNewLabel_Useranme;
	JLabel lblNewLabel_Password;

	public static JCheckBox checkBoxSplitPst;
	public static JSpinner spinner_MB;
	public static JSpinner spinner_GB;

	public static JComboBox comboBoxNamingConvention;
	public static JCheckBox chckbxNamingconvention;
	private final ButtonGroup buttonGroup_1 = new ButtonGroup();
	public static JRadioButton r_emlx;
	public static JRadioButton r_mbox;
	public static JRadioButton r_html;
	public static int version = 0;
	private final ButtonGroup buttonGroup_2 = new ButtonGroup();
	public static JTextField outputUsernameField;
	public static JPasswordField passwordField;
	private JScrollPane scrollPane_2;
	JPasswordField passwordField_1;
	static EmailWizardApplication frame;
	JPanel DateFilterPanel;
	private JPanel MsOfficePanel_P6;
	public static JRadioButton r_mailbox;
	public static JRadioButton r_public;
	public static JRadioButton r_archive;
	private final ButtonGroup buttonGroup_3 = new ButtonGroup();
	private JButton b_next1;
	private JButton b_back1;
	private JPanel p_outputEmailLogin;

	ArrayList<String> EMAILCLIENTLIST;
	public static JRadioButton r_Eml;
	public static JRadioButton r_pdf;
	public static JRadioButton r_pst;
	public static JRadioButton r_msg;
	public static JRadioButton r_office;
	public static JRadioButton r_yandex;
	public static JRadioButton r_gmail;
	public static JRadioButton r_yahoo;
	public static JRadioButton r_aol;
	public static JRadioButton r_zoho;
	public static JRadioButton r_rtf;
	public static JRadioButton r_xps;

	public static JRadioButton r_emf;
	public static JRadioButton r_docx;
	public static JRadioButton r_jpeg;
	public static JRadioButton r_docm;
	public static JRadioButton r_text;
	public static JRadioButton r_tiff;
	public static JRadioButton r_png;
	public static JRadioButton r_svg;
	public static JRadioButton r_epub;
	public static JRadioButton r_dotm;
	public static JRadioButton r_ott;
	public static JRadioButton r_gif;
	public static JRadioButton r_bmp;
	public static JRadioButton r_wordml;
	public static JRadioButton r_odt;
	public static JRadioButton r_csv;
	public static JRadioButton r_imap;
	public static JRadioButton r_hostgator;
	public static JRadioButton r_hotmail;
	public static JRadioButton r_aws;
	public static JRadioButton r_icloud;

	public String outputSource = OutputSource.EML.name();
	private JTextField txtCloudhostgatorcom;
	private JTextField textField_portOutput;
	JLabel lbl_Hostoutput;
	JLabel lblNewLabel;

	DefaultMutableTreeNode root;
	DefaultMutableTreeNode subNode;
	CheckboxTree folderTree;
	public static List<String> folderIdlist = new ArrayList<>();

	public static JCheckBox chckbxSaveSeperateAttachments;
	public static org.slf4j.Logger logger = LoggerFactory.getLogger(EmailWizardApplication.class);

	public final static int IMAP_RERESH_TIMEOUT = 240000;
	public final static int DEMO_LIMIT = 15;
	public final static int IMAP_MAIL_SIZE = 25000000;
	private JPanel panel_LogScreen;
	public static JTextPane textPane_log;
	private JButton btnNewButton;
	public static File pstSplitFile;
	public static PersonalStorage pst;
	public static FolderInfo pstfolderInfo;
	public static int splitCount = 0;

	public static JCheckBox chckbxSkip_body;
	public static JCheckBox chckbxSkip_subject;
	public static JCheckBox chckbxSkip_date;
	public static JCheckBox chckbxSkip_from;
	private JCheckBox chckbx_Proxy;
	private JPanel panel_TableLogin;
	public static JTable table_Login;
	public static DefaultTableModel loginTableModel;
	private JButton btn_brwCSV;
	private JTextField textField_brwCSV;
	private JLabel lblNewLabel_2;
	public static JRadioButton r_gmail_app;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					try {
						UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
						frame = new EmailWizardApplication(true, 0);
						frame.setLocationRelativeTo(null);
						frame.setResizable(false);
						frame.setVisible(true);

					} catch (ClassNotFoundException e1) {

					} catch (InstantiationException e1) {

					} catch (IllegalAccessException e1) {

					} catch (UnsupportedLookAndFeelException e1) {

					}

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public EmailWizardApplication(boolean demoCheck, int versiontype) {

		logger.info("!!!!!!!!!---* Email Wizard Application Started *----!!!!!!!!!!");
		version = versiontype;
		this.demo = demoCheck;
		AsposeActivation lic = new AsposeActivation();
		lic.doAsposeLicActivation();
		if (demo) {
			setTitle(ToolDetails.messageboxtitle + " Demo");
		} else {
			setTitle(ToolDetails.messageboxtitle + " Full");

		}

		addWindowListener(new WindowAdapter() {

			public void windowClosing(WindowEvent arg0) {

				String warn = "Do you want to close the application?";
				int ans = JOptionPane.showConfirmDialog(EmailWizardApplication.this, warn, ToolDetails.messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(EmailWizardApplication.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					openBrowser(ToolDetails.infourl);
					System.exit(0);
				}

			}
		});

		setIconImage(Toolkit.getDefaultToolkit().getImage(EmailWizardApplication.class.getResource("/128x128.png")));
		setLocationRelativeTo(null);
		setResizable(false);
		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		setBounds(100, 100, 796, 488);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		CardLayout = new JPanel();
		CardLayout.setBounds(0, 48, 780, 408);
		contentPane.add(CardLayout);
		CardLayout.setLayout(new CardLayout(0, 0));

		JPanel LoginPanel_P1 = new JPanel();
		LoginPanel_P1.setBackground(Color.WHITE);
		CardLayout.add(LoginPanel_P1, "GoogleLoginPanel_1");
		LoginPanel_P1.setLayout(null);

		LoginPanel_1 = new JPanel();
		LoginPanel_1.setBounds(165, 0, 615, 362);
		LoginPanel_1.setBorder(new TitledBorder(new EtchedBorder(EtchedBorder.LOWERED, null, null), "",
				TitledBorder.LEADING, TitledBorder.TOP, null, null));
		LoginPanel_1.setBackground(Color.WHITE);
		LoginPanel_P1.add(LoginPanel_1);
		LoginPanel_1.setLayout(new CardLayout(0, 0));

		JPanel panel_login = new JPanel();
		panel_login.setBackground(Color.WHITE);
		panel_login.setBorder(new TitledBorder(new MatteBorder(2, 2, 2, 2, (Color) new Color(0, 0, 255)), "",
				TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 255)));
		LoginPanel_1.add(panel_login, "panel_login");
		panel_login.setLayout(null);

		chckbx_Proxy = new JCheckBox("IMAP Host & Port No. Settings");
		chckbx_Proxy.setBounds(88, 193, 199, 23);
		panel_login.add(chckbx_Proxy);
		chckbx_Proxy.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == ItemEvent.SELECTED) {

					lblServiceaccountIdAndHostName.setText("Host Name");
					lblServiceaccountIdAndHostName.setVisible(true);
					txtServiceAccountIDorImapHostName.setVisible(true);

					lblPFileAndPortNumber.setText("Port No.");
					textField_p12FileAndPortNo.setVisible(true);
					lblPFileAndPortNumber.setVisible(true);

				} else if (e.getStateChange() == ItemEvent.DESELECTED) {

					lblServiceaccountIdAndHostName.setText("Host Name");
					lblServiceaccountIdAndHostName.setVisible(false);
					txtServiceAccountIDorImapHostName.setVisible(false);

					lblPFileAndPortNumber.setText("Port No.");
					textField_p12FileAndPortNo.setVisible(false);
					lblPFileAndPortNumber.setVisible(false);

				}
			}
		});
		chckbx_Proxy.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbx_Proxy.setBackground(Color.WHITE);

		btn_Login = new JButton("");
		btn_Login.setBounds(269, 272, 127, 39);
		panel_login.add(btn_Login);
		btn_Login.setRolloverEnabled(false);
		btn_Login.setRequestFocusEnabled(false);
		btn_Login.setOpaque(false);
		btn_Login.setFocusable(false);
		btn_Login.setFocusTraversalKeysEnabled(false);
		btn_Login.setFocusPainted(false);
		btn_Login.setDefaultCapable(false);
		btn_Login.setContentAreaFilled(false);
		btn_Login.setBorderPainted(false);
		btn_Login.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btn_Login.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/sign-in-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btn_Login.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_Login.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/sign-in-btn.png")));

		btnBrowseP2File = new JButton("");
		btnBrowseP2File.setBounds(358, 188, 124, 28);
		panel_login.add(btnBrowseP2File);
		btnBrowseP2File.setVisible(false);
		btnBrowseP2File.setRolloverEnabled(false);
		btnBrowseP2File.setRequestFocusEnabled(false);
		btnBrowseP2File.setOpaque(false);
		btnBrowseP2File.setFocusable(false);
		btnBrowseP2File.setFocusTraversalKeysEnabled(false);
		btnBrowseP2File.setFocusPainted(false);
		btnBrowseP2File.setDefaultCapable(false);
		btnBrowseP2File.setContentAreaFilled(false);
		btnBrowseP2File.setBorderPainted(false);
		btnBrowseP2File.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnBrowseP2File.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/browse-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnBrowseP2File.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/browse-btn.png")));
			}
		});

		btnBrowseP2File.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/browse-btn.png")));
		btnBrowseP2File.setBackground(Color.WHITE);

		lblPFileAndPortNumber = new JLabel("p12 File");
		lblPFileAndPortNumber.setBounds(93, 167, 72, 14);
		panel_login.add(lblPFileAndPortNumber);
		lblPFileAndPortNumber.setVisible(false);
		lblPFileAndPortNumber.setFont(new Font("Tahoma", Font.BOLD, 13));

		lblServiceaccountIdAndHostName = new JLabel("Service Account ID");
		lblServiceaccountIdAndHostName.setBounds(92, 134, 127, 14);
		panel_login.add(lblServiceaccountIdAndHostName);
		lblServiceaccountIdAndHostName.setVisible(false);
		lblServiceaccountIdAndHostName.setFont(new Font("Tahoma", Font.BOLD, 13));

		JLabel lblUsername = new JLabel("");
		lblUsername.setBounds(309, 32, 32, 32);
		panel_login.add(lblUsername);
		lblUsername.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/username.png")));

		lblUserName = new JLabel("User Name");
		lblUserName.setBounds(95, 64, 69, 23);
		panel_login.add(lblUserName);
		lblUserName.setFont(new Font("Tahoma", Font.BOLD, 13));

		InputtxtUserName = new JTextField();
		InputtxtUserName.setBounds(230, 66, 244, 20);
		panel_login.add(InputtxtUserName);
		InputtxtUserName.setText("ajay1khanduri@gmail.com");
		TextFieldPopup(InputtxtUserName);
		InputtxtUserName.setColumns(10);

		lblPassworduser = new JLabel("Password");
		lblPassworduser.setBounds(92, 100, 79, 14);

		panel_login.add(lblPassworduser);
		lblPassworduser.setFont(new Font("Tahoma", Font.BOLD, 13));

		passwordField_1 = new JPasswordField();
		passwordField_1.setBounds(230, 98, 243, 20);
		panel_login.add(passwordField_1);

		txtServiceAccountIDorImapHostName = new JTextField();
		txtServiceAccountIDorImapHostName.setBounds(231, 132, 243, 20);
		TextFieldPopup(txtServiceAccountIDorImapHostName);
		panel_login.add(txtServiceAccountIDorImapHostName);
		// txtServiceAccountIDorImapHostName.setText(inputIMAPHostName);

		txtServiceAccountIDorImapHostName.setVisible(false);
		txtServiceAccountIDorImapHostName.setColumns(10);

		textField_p12FileAndPortNo = new JTextField();
		textField_p12FileAndPortNo.setBounds(231, 165, 243, 20);
		TextFieldPopup(textField_p12FileAndPortNo);
		panel_login.add(textField_p12FileAndPortNo);
		textField_p12FileAndPortNo.setVisible(false);
		textField_p12FileAndPortNo.setColumns(10);

		label_LoginGif = new JLabel("");
		label_LoginGif.setBounds(300, 234, 46, 32);
		panel_login.add(label_LoginGif);
		label_LoginGif.setVisible(false);
		label_LoginGif.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/loading.gif")));

		panel_TableLogin = new JPanel();
		panel_TableLogin.setBackground(Color.WHITE);
		LoginPanel_1.add(panel_TableLogin, "panel_TableLogin");
		panel_TableLogin.setLayout(null);

		JScrollPane scrollPane_5 = new JScrollPane();
		scrollPane_5.setBounds(0, 0, 611, 186);
		panel_TableLogin.add(scrollPane_5);

		table_Login = new JTable();
		table_Login.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "S.No", "Email ID", "Password", "Imap Host", "Port No", "Status" }));
		table_Login.setRowHeight(table_Login.getRowHeight() + 10);
		table_Login.setEnabled(false);
		scrollPane_5.setViewportView(table_Login);

		JButton btn_TableLogin = new JButton("");
		btn_TableLogin.setBounds(480, 321, 131, 38);
		btn_TableLogin.setRolloverEnabled(false);
		btn_TableLogin.setRequestFocusEnabled(false);
		btn_TableLogin.setOpaque(false);
		btn_TableLogin.setFocusable(false);
		btn_TableLogin.setFocusTraversalKeysEnabled(false);
		btn_TableLogin.setFocusPainted(false);
		btn_TableLogin.setDefaultCapable(false);
		btn_TableLogin.setContentAreaFilled(false);
		btn_TableLogin.setBorderPainted(false);
		btn_TableLogin.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btn_TableLogin.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btn_TableLogin.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
			}
		});

		btn_TableLogin.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
		btn_TableLogin.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				int rows = table_Login.getRowCount();
				if (rows > 0) {

					changeHeader();
					changeHeaderoutput();
					// buttonDisables();

					if (selectedInput.equals(InputSource.IMAP.getValue())
							|| selectedInput.equals(InputSource.HOSTGATOR.getValue()) || chckbx_Proxy.isSelected()) {
						inputIMAPHostName = txtServiceAccountIDorImapHostName.getText().trim();
						inputIMAPPortNo = Integer.parseInt(textField_p12FileAndPortNo.getText().trim());
					}
					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleDownloadOptions_3");
				} else {

					LogUtils.setTextToLogScreen(textPane_log, logger,
							"Please Add Email Account In Table For Backup!!!");

					JOptionPane.showMessageDialog(EmailWizardApplication.this,
							"Please Add Email Account In Table For Backup!", ToolDetails.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
				}

			}
		});
		panel_TableLogin.add(btn_TableLogin);

		btn_brwCSV = new JButton("");
		btn_brwCSV.setBounds(493, 186, 116, 38);
		btn_brwCSV.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				JFileChooser jFileChooser = new JFileChooser();
				jFileChooser.setBackground(Color.WHITE);
				jFileChooser.setMultiSelectionEnabled(false);
				jFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				FileNameExtensionFilter filter = new FileNameExtensionFilter(".csv", "CSV");
				jFileChooser.setFileFilter(filter);
				jFileChooser.setAcceptAllFileFilterUsed(false);
				jFileChooser.addChoosableFileFilter(filter);
				if (jFileChooser.showOpenDialog(EmailWizardApplication.this) == JFileChooser.APPROVE_OPTION) {
					File file = jFileChooser.getSelectedFile();
					if (!(file == null)) {
						try {
							textField_brwCSV.setText(file.getAbsolutePath());
							loginTableModel = (DefaultTableModel) EmailWizardApplication.table_Login.getModel();
							CSVUtils.readCSV(loginTableModel, textField_brwCSV.getText().trim());
						} catch (Exception ex) {
							// TODO Auto-generated catch block
							ex.printStackTrace();
						}

					}

				}

			}
		});
		btn_brwCSV.setRolloverEnabled(false);
		btn_brwCSV.setRequestFocusEnabled(false);
		btn_brwCSV.setOpaque(false);
		btn_brwCSV.setFocusable(false);
		btn_brwCSV.setFocusTraversalKeysEnabled(false);
		btn_brwCSV.setFocusPainted(false);
		btn_brwCSV.setDefaultCapable(false);
		btn_brwCSV.setContentAreaFilled(false);
		btn_brwCSV.setBorderPainted(false);
		btn_brwCSV.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btn_brwCSV.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btn_brwCSV.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
			}
		});

		btn_brwCSV.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
		panel_TableLogin.add(btn_brwCSV);

		textField_brwCSV = new JTextField();
		textField_brwCSV.setBounds(4, 191, 489, 27);
		textField_brwCSV.setEditable(false);
		panel_TableLogin.add(textField_brwCSV);
		textField_brwCSV.setColumns(10);

		JScrollPane scrollPane_6 = new JScrollPane();
		scrollPane_6.setBounds(0, 229, 601, 86);
		panel_TableLogin.add(scrollPane_6);

		JTextArea txtraddLoginDetails = new JTextArea();
		scrollPane_6.setViewportView(txtraddLoginDetails);
		txtraddLoginDetails.setForeground(new Color(255, 0, 0));
		txtraddLoginDetails.setFont(new Font("Californian FB", Font.BOLD, 13));
		txtraddLoginDetails.setLineWrap(true);
		txtraddLoginDetails.setText(
				"               (1). Click browse CSV file to add login details in the above table\r\n               (2). Make sure that 1 row of your CSV file should be like this Ex:-\r\n                       1.|Email ID | Password | IMAP Host | Port No.|   \r\n               (3). Start adding your login details from the 2 row Ex:-\r\n                       2.|xyz@gmail.com|  xzy123 |imap.gmail.com  | 993 |\r\n               (4). For Imap Host and Port No. check your email client setting Ex:-\r\n                       Gmail Imap Host : imap.gmail.com\r\n                       Gmail Port   No.   : 993");

		lblNewLabel_2 = new JLabel("Click here to download CSV file sample");

		lblNewLabel_2.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
		lblNewLabel_2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				if (Desktop.isDesktopSupported()) {
					Desktop desktop = Desktop.getDesktop();

					try {
						Workbook workbook = CSVUtils.createSampleCSVStructure();
						File desktopPath = new File(System.getProperty("user.home"),
								"Desktop" + File.separator + "CSVCredentialSample" + ".csv");
						workbook.save(desktopPath.getAbsolutePath(), com.aspose.cells.SaveFormat.CSV);
						workbook.dispose();
						desktop.open(desktopPath);
					} catch (IOException | URISyntaxException ex) {
						logger.warn("Warning : " + ex.getMessage());
					} catch (Exception e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			}
		});
		lblNewLabel_2.setForeground(new Color(0, 0, 153));
		lblNewLabel_2.setFont(new Font("Calibri", Font.BOLD, 14));
		lblNewLabel_2.setBounds(96, 321, 251, 18);
		Font font = lblNewLabel_2.getFont();
		Map<TextAttribute, Object> attributes = new HashMap<>(font.getAttributes());
		attributes.put(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON);
		lblNewLabel_2.setFont(font.deriveFont(attributes));
		panel_TableLogin.add(lblNewLabel_2);

		btnBrowseP2File.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				JFileChooser jFileChooser = new JFileChooser(
						System.getProperty("user.home") + File.separator + "Desktop");

				jFileChooser.setBackground(Color.WHITE);

				jFileChooser.setAcceptAllFileFilterUsed(false);
				FileNameExtensionFilter filter = new FileNameExtensionFilter(".p12", "P12");
				jFileChooser.addChoosableFileFilter(filter);
				if (jFileChooser.showOpenDialog(EmailWizardApplication.this) == JFileChooser.APPROVE_OPTION) {
					File file = jFileChooser.getSelectedFile();

					textField_p12FileAndPortNo.setText(file.getAbsolutePath());

				}

			}
		});
		btn_Login.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				Thread loginThread = new Thread(new Runnable() {

					@Override
					public void run() {

						changeHeader();
						changeHeaderoutput();
						buttonDisables();
						if (selectedInput.equals(InputSource.GSUITE.getValue())) {
							GSuiteLogin();
						} else if (selectedInput.equals(InputSource.GMAIL_APP.getValue())) {
							InputGmailAPPLogin();
						} else if (selectedInput.equals(InputSource.MS_Office_365.getValue())) {
							EWSLogin();
						} else {
							if (selectedInput.equals(InputSource.IMAP.getValue())
									|| selectedInput.equals(InputSource.HOSTGATOR.getValue())
									|| chckbx_Proxy.isSelected()) {
								inputIMAPHostName = txtServiceAccountIDorImapHostName.getText().trim();
								inputIMAPPortNo = Integer.parseInt(textField_p12FileAndPortNo.getText().trim());
							}
							IMAPLogin();
						}

					}
				});
				loginThread.start();

			}
		});

		l_bottombar = new JLabel("");
		l_bottombar.setBounds(0, 362, 780, 38);
		l_bottombar.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/bottomn.png")));

		LoginPanel_P1.add(l_bottombar);
		scrollPane_2 = new JScrollPane();
		scrollPane_2.setBounds(0, 0, 166, 362);
		LoginPanel_P1.add(scrollPane_2);

		@SuppressWarnings("unchecked")
		JList sourceInput_list = new JList(InputSource.getDefaultListModel());
		sourceInput_list.setSelectedIndex(0);
		sourceInput_list.setFont(new Font("Tahoma", Font.BOLD, 13));
		sourceInput_list.addMouseListener(new MouseAdapter() {
			@Override
			public void mousePressed(MouseEvent e) {

				InputSource is = (InputSource) sourceInput_list.getSelectedValue();
				selectedInput = is.getValue();
				selectInputSource(selectedInput);

			}
		});

		scrollPane_2.setViewportView(sourceInput_list);

		JPanel TreeStructurePanel_P2 = new JPanel();
		CardLayout.add(TreeStructurePanel_P2, "treePanel");
		TreeStructurePanel_P2.setLayout(null);

		JScrollPane scrollPane_3 = new JScrollPane();
		scrollPane_3.setBounds(0, 0, 780, 362);
		TreeStructurePanel_P2.add(scrollPane_3);

		folderTree = new CheckboxTree();
		scrollPane_3.setViewportView(folderTree);
		folderTree.setModel(new DefaultTreeModel(new DefaultMutableTreeNode("JTree") {
			{
			}
		}));

		JButton btnTreeBack = new JButton("Back");
		btnTreeBack.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				DefaultTreeModel model = (DefaultTreeModel) folderTree.getModel();
				DefaultMutableTreeNode root = (DefaultMutableTreeNode) model.getRoot();
				root.removeAllChildren();

				if (InputSource.MS_Office_365.getValue().equals(selectedInput)) {

					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "p_msoffice");

				} else {
					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleLoginPanel_1");
				}
			}
		});
		btnTreeBack.setBounds(10, 373, 89, 23);
		TreeStructurePanel_P2.add(btnTreeBack);

		JButton btnTreeNext = new JButton("Next");
		btnTreeNext.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				TreePath[] treePath = folderTree.getCheckingPaths();

				if (treePath.length == 0) {
					JOptionPane.showMessageDialog(frame, "Select File From the Tree", ToolDetails.messageboxtitle,
							JOptionPane.ERROR_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information-2.png")));

				} else {

					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleDownloadOptions_3");
				}

			}
		});
		btnTreeNext.setBounds(663, 373, 89, 23);
		TreeStructurePanel_P2.add(btnTreeNext);

		JPanel UserDetailsAndFolderPanel_P3 = new JPanel();
		UserDetailsAndFolderPanel_P3.setBackground(Color.WHITE);
		CardLayout.add(UserDetailsAndFolderPanel_P3, "GoogleUserDetailsPanel_2");
		UserDetailsAndFolderPanel_P3.setLayout(null);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(0, 0, 780, 351);
		UserDetailsAndFolderPanel_P3.add(scrollPane);

		table_UserDetails = new JTable() {
			@Override
			public boolean isCellEditable(int row, int column) {

				return column == 3;

			};

			@Override
			public Point getToolTipLocation(MouseEvent event) {
				return new Point(10, 10);
			}

		};
		table_UserDetails.setRowHeight(table_UserDetails.getRowHeight() + 10);
		table_UserDetails.setDefaultRenderer(Object.class, new CellRenderer());

		table_UserDetails.addMouseMotionListener(new MouseMotionListener() {
			@Override
			public void mouseDragged(MouseEvent e) {

			}

			@Override
			public void mouseMoved(MouseEvent e) {

				int row = table_UserDetails.rowAtPoint(e.getPoint());
				if (row > -1) {

					table_UserDetails.clearSelection();
					table_UserDetails.setRowSelectionInterval(row, row);
				} else {
					table_UserDetails.setSelectionBackground(Color.blue);
				}

			}
		});

		DefaultTableModel model = new DefaultTableModel() {
			public Class<?> getColumnClass(int column) {
				switch (column) {
				case 0:
					return String.class;
				case 1:
					return String.class;
				case 2:
					return String.class;
				case 3:
					return Boolean.class;
				default:
					return String.class;
				}
			}
		};

		// ASSIGN THE MODEL TO TABLE
		table_UserDetails.getTableHeader().setReorderingAllowed(false);
		table_UserDetails.setModel(model);

		model.addColumn("<HTML><B>S.No</B></HTML>");
		model.addColumn("<HTML><B>Folders Name</B></HTML>");
		model.addColumn("<HTML><B>Count</B></HTML>");
		model.addColumn(Status.INDETERMINATE);

		model.addTableModelListener(new HeaderCheckBoxHandler(table_UserDetails));

		TableCellRenderer r = new HeaderRenderer(table_UserDetails.getTableHeader(), 3);
		table_UserDetails.getColumnModel().getColumn(3).setHeaderRenderer(r);
		table_UserDetails.getTableHeader().setReorderingAllowed(false);

		scrollPane.setViewportView(table_UserDetails);

		btnNext_p2 = new JButton("");
		btnNext_p2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				int rows = table_UserDetails.getRowCount();
				int count = 0;
				for (int i = 0; i < rows; i++) {

					Object checked = table_UserDetails.getValueAt(i, 3);
					if (!(boolean) checked) {
						count++;
					}
				}
				if (count != rows) {

					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleDownloadOptions_3");
				} else {

					LogUtils.setTextToLogScreen(textPane_log, logger, "Please select atleast One User!!!");

					JOptionPane.showMessageDialog(EmailWizardApplication.this, "Please select atleast One User!",
							ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
				}

			}
		});
		btnNext_p2.setBounds(663, 356, 107, 34);
		btnNext_p2.setRolloverEnabled(false);
		btnNext_p2.setRequestFocusEnabled(false);
		btnNext_p2.setOpaque(false);
		btnNext_p2.setFocusable(false);
		btnNext_p2.setFocusTraversalKeysEnabled(false);
		btnNext_p2.setFocusPainted(false);
		btnNext_p2.setDefaultCapable(false);
		btnNext_p2.setContentAreaFilled(false);
		btnNext_p2.setBorderPainted(false);
		btnNext_p2.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnNext_p2.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnNext_p2.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
			}
		});

		btnNext_p2.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));

		UserDetailsAndFolderPanel_P3.add(btnNext_p2);

		btnBack_p2 = new JButton("");
		btnBack_p2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				DefaultTableModel dm = (DefaultTableModel) table_UserDetails.getModel();
				while (dm.getRowCount() > 0) {
					dm.removeRow(0);
				}
				count = 0;
				if (InputSource.MS_Office_365.getValue().equals(selectedInput)) {

					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "p_msoffice");

				} else {
					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleLoginPanel_1");
				}

			}
		});
		btnBack_p2.setBounds(0, 358, 112, 34);

		btnBack_p2.setRolloverEnabled(false);
		btnBack_p2.setRequestFocusEnabled(false);
		btnBack_p2.setOpaque(false);
		btnBack_p2.setFocusable(false);
		btnBack_p2.setFocusTraversalKeysEnabled(false);
		btnBack_p2.setFocusPainted(false);
		btnBack_p2.setDefaultCapable(false);
		btnBack_p2.setContentAreaFilled(false);
		btnBack_p2.setBorderPainted(false);
		btnBack_p2.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnBack_p2.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnBack_p2.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
			}
		});

		btnBack_p2.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));

		UserDetailsAndFolderPanel_P3.add(btnBack_p2);

		JLabel label_7 = new JLabel("");
		label_7.setBounds(0, 352, 780, 46);
		label_7.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/bottomn.png")));

		UserDetailsAndFolderPanel_P3.add(label_7);

		JPanel SavingFormatPanel_P4 = new JPanel();
		SavingFormatPanel_P4.setBackground(Color.WHITE);
		CardLayout.add(SavingFormatPanel_P4, "GoogleDownloadOptions_3");
		SavingFormatPanel_P4.setLayout(null);

		btnNew_p3 = new JButton("");
		btnNew_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				LogUtils.setTextToLogScreen(textPane_log, logger, "Selected Output Source  : " + outputSource);

				if (c_contact.isSelected() || c_drive.isSelected() || c_photos.isSelected() || c_calendar.isSelected()
						|| c_email.isSelected()) {

					if (rdbtnDateFilter.isSelected()) {

						Calendar calendarstartdate = EmailWizardApplication.start_dateChooser.getCalendar();
						Calendar calendarenddate = EmailWizardApplication.end_dateChooser.getCalendar();

						if (calendarstartdate != null && calendarenddate != null) {

							calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
							calendarstartdate.set(Calendar.MINUTE, 00);
							calendarstartdate.set(Calendar.SECOND, 00);

							calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
							calendarenddate.set(Calendar.MINUTE, 59);
							calendarenddate.set(Calendar.SECOND, 59);

							Long startDateMillisecond = calendarstartdate.getTimeInMillis();
							Long endateMillisecond = calendarenddate.getTimeInMillis();

							if (startDateMillisecond <= endateMillisecond) {
								CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
								card.show(EmailWizardApplication.CardLayout, "GoogleDownloading_4");
							} else {
								LogUtils.setTextToLogScreen(textPane_log, logger,
										"End date cannot be smaller than start date");

								JOptionPane.showMessageDialog(EmailWizardApplication.this,
										"Please enter the correct date end date can not be smaller than start date.",
										ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
										new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
							}

						} else {
							LogUtils.setTextToLogScreen(textPane_log, logger, "Please Select Date");

							JOptionPane.showMessageDialog(EmailWizardApplication.this,
									"Please Select Start date and End date from the calendar.",
									ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
						}
					} else {
						CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
						card.show(EmailWizardApplication.CardLayout, "GoogleDownloading_4");
					}

				} else {
					LogUtils.setTextToLogScreen(textPane_log, logger, "Please select atleast One option!!!");

					JOptionPane.showMessageDialog(EmailWizardApplication.this, "Please select atleast One option!!!",
							ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

				}

			}
		});
		btnNew_p3.setBounds(648, 359, 124, 34);

		btnNew_p3.setRolloverEnabled(false);
		btnNew_p3.setRequestFocusEnabled(false);
		btnNew_p3.setOpaque(false);
		btnNew_p3.setFocusable(false);
		btnNew_p3.setFocusTraversalKeysEnabled(false);
		btnNew_p3.setFocusPainted(false);
		btnNew_p3.setDefaultCapable(false);
		btnNew_p3.setContentAreaFilled(false);
		btnNew_p3.setBorderPainted(false);
		btnNew_p3.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnNew_p3.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnNew_p3.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
			}
		});

		btnNew_p3.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));

		SavingFormatPanel_P4.add(btnNew_p3);

		btnBack_p3 = new JButton("");
		btnBack_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				if (InputSource.GMAIL_APP.getValue().equals(selectedInput)) {
					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleLoginPanel_1");

				} else if (InputSource.Bulk.getValue().equals(selectedInput)) {
					DefaultTableModel dm = (DefaultTableModel) table_Login.getModel();

					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleLoginPanel_1");

					CardLayout loginCardLayout = (CardLayout) LoginPanel_1.getLayout();
					loginCardLayout.show(LoginPanel_1, "panel_TableLogin");
				}

				else {

					CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
					card.show(EmailWizardApplication.CardLayout, "GoogleUserDetailsPanel_2");

				}
				checkBoxSplitPst.setSelected(false);
				chckbxSkipDuplicate.setSelected(false);
				chckbxNamingconvention.setSelected(false);

			}
		});
		btnBack_p3.setBounds(5, 358, 119, 34);
		btnBack_p3.setRolloverEnabled(false);
		btnBack_p3.setRequestFocusEnabled(false);
		btnBack_p3.setOpaque(false);
		btnBack_p3.setFocusable(false);
		btnBack_p3.setFocusTraversalKeysEnabled(false);
		btnBack_p3.setFocusPainted(false);
		btnBack_p3.setDefaultCapable(false);
		btnBack_p3.setContentAreaFilled(false);
		btnBack_p3.setBorderPainted(false);
		btnBack_p3.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnBack_p3.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnBack_p3.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
			}
		});

		btnBack_p3.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));

		SavingFormatPanel_P4.add(btnBack_p3);

		SavingOptionPanel = new JPanel();
		SavingOptionPanel.setForeground(Color.BLUE);

		SavingOptionPanel.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)), "",
				TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 255)));
		SavingOptionPanel.setBackground(Color.WHITE);
		SavingOptionPanel.setBounds(0, 0, 780, 356);
		SavingFormatPanel_P4.add(SavingOptionPanel);
		SavingOptionPanel.setLayout(null);

		SavingOptionPanel_1 = new JPanel();
		SavingOptionPanel_1.setVisible(true);
		SavingOptionPanel_1.setBorder(

				new TitledBorder(
						new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)), "",
						TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 255)));
		SavingOptionPanel_1.setBackground(Color.WHITE);
		SavingOptionPanel_1.setBounds(0, 68, 782, 288);
		SavingOptionPanel.add(SavingOptionPanel_1);

		SavingOptionPanel_1.setLayout(new CardLayout(0, 0));

		p_outputSavingformat = new JPanel();
		p_outputSavingformat.setBackground(new Color(255, 255, 255));
		SavingOptionPanel_1.add(p_outputSavingformat, "p_outputSavingformat");
		p_outputSavingformat.setLayout(null);

		r_Eml = new JRadioButton("EML");
		r_Eml.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fileSavingFormatEvent();
				outputSource = OutputSource.EML.name();

			}
		});
		r_Eml.setBounds(6, 7, 56, 23);
		p_outputSavingformat.add(r_Eml);
		r_Eml.setSelected(true);
		buttonGroup.add(r_Eml);
		r_Eml.setBackground(Color.WHITE);

		r_pdf = new JRadioButton("PDF");
		r_pdf.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fileSavingFormatEvent();
				outputSource = OutputSource.PDF.name();
			}
		});
		r_pdf.setBounds(117, 7, 56, 23);
		p_outputSavingformat.add(r_pdf);
		buttonGroup.add(r_pdf);
		r_pdf.setBackground(Color.WHITE);

		r_pst = new JRadioButton("PST");
		r_pst.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (r_pst.isSelected()) {
					chckbxNamingconvention.setVisible(false);
					comboBoxNamingConvention.setVisible(false);
					checkBoxSplitPst.setVisible(true);
					radioButtonMB.setVisible(false);
					rdbtnGb.setVisible(false);
					spinner_GB.setVisible(false);
					spinner_MB.setVisible(false);
					chckbxSaveSeperateAttachments.setVisible(false);
				}
				outputSource = OutputSource.PST.name();

			}
		});
		buttonGroup.add(r_pst);
		r_pst.setBackground(Color.WHITE);
		r_pst.setBounds(68, 7, 47, 23);
		p_outputSavingformat.add(r_pst);

		r_msg = new JRadioButton("MSG");
		r_msg.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.MSG.name();
			}
		});
		buttonGroup.add(r_msg);
		r_msg.setBackground(Color.WHITE);
		r_msg.setBounds(175, 7, 56, 23);
		p_outputSavingformat.add(r_msg);

		r_emlx = new JRadioButton("EMLX");
		r_emlx.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.EMLX.name();
			}
		});
		buttonGroup.add(r_emlx);
		r_emlx.setBackground(Color.WHITE);
		r_emlx.setBounds(231, 7, 63, 23);
		p_outputSavingformat.add(r_emlx);

		r_mbox = new JRadioButton("MBOX");
		buttonGroup.add(r_mbox);
		r_mbox.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fileSavingFormatEvent();
				outputSource = OutputSource.MBOX.name();
			}
		});
		r_mbox.setBackground(Color.WHITE);
		r_mbox.setBounds(296, 7, 63, 23);
		p_outputSavingformat.add(r_mbox);

		r_html = new JRadioButton("HTML");
		buttonGroup.add(r_html);
		r_html.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fileSavingFormatEvent();
				outputSource = OutputSource.HTML.name();
			}
		});

		r_html.setBackground(Color.WHITE);
		r_html.setBounds(368, 7, 56, 23);
		p_outputSavingformat.add(r_html);

		r_office = new JRadioButton("Office365");
		buttonGroup.add(r_office);
		r_office.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.Office365.name();
			}
		});
		r_office.setBackground(Color.WHITE);
		r_office.setBounds(6, 186, 80, 23);
		p_outputSavingformat.add(r_office);

		r_gmail = new JRadioButton("Gmail");
		r_gmail.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.GMAIL.name();
			}
		});
		buttonGroup.add(r_gmail);
		r_gmail.setBackground(Color.WHITE);
		r_gmail.setBounds(6, 160, 56, 23);
		p_outputSavingformat.add(r_gmail);

		r_yahoo = new JRadioButton("Yahoo");
		buttonGroup.add(r_yahoo);
		r_yahoo.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.YAHOO.name();
			}
		});
		r_yahoo.setBackground(Color.WHITE);
		r_yahoo.setBounds(96, 160, 63, 23);
		p_outputSavingformat.add(r_yahoo);

		r_aol = new JRadioButton("Aol");
		buttonGroup.add(r_aol);
		r_aol.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.AOL.name();
			}
		});
		r_aol.setBackground(Color.WHITE);
		r_aol.setBounds(6, 212, 47, 23);
		p_outputSavingformat.add(r_aol);

		r_zoho = new JRadioButton("Zoho");
		buttonGroup.add(r_zoho);
		r_zoho.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.ZOHO_EMAIL.name();
			}
		});
		r_zoho.setBackground(Color.WHITE);
		r_zoho.setBounds(96, 212, 56, 23);
		p_outputSavingformat.add(r_zoho);

		r_yandex = new JRadioButton("Yandex");
		buttonGroup.add(r_yandex);
		r_yandex.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.YANDEX.name();
			}
		});
		r_yandex.setBackground(Color.WHITE);
		r_yandex.setBounds(96, 186, 63, 23);
		p_outputSavingformat.add(r_yandex);

		p_outputEmailLogin = new JPanel();
		p_outputEmailLogin.setBorder(null);
		p_outputEmailLogin.setBackground(new Color(255, 255, 255));
		p_outputEmailLogin.setBounds(504, 7, 267, 266);
		p_outputSavingformat.add(p_outputEmailLogin);
		p_outputEmailLogin.setLayout(null);

		lblNewLabel_Useranme = new JLabel("UserName");
		lblNewLabel_Useranme.setBackground(SystemColor.activeCaption);
		lblNewLabel_Useranme.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_Useranme.setBounds(87, 22, 67, 14);
		p_outputEmailLogin.add(lblNewLabel_Useranme);

		btn_login = new JButton("");
		btn_login.setBounds(64, 215, 143, 39);
		btn_login.setRolloverEnabled(false);
		btn_login.setRequestFocusEnabled(false);
		btn_login.setOpaque(false);
		btn_login.setFocusable(false);
		btn_login.setFocusTraversalKeysEnabled(false);
		btn_login.setFocusPainted(false);
		btn_login.setDefaultCapable(false);
		btn_login.setContentAreaFilled(false);
		btn_login.setBorderPainted(false);
		btn_login.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btn_login.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/sign-in-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btn_login.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/sign-in-hvr-btn.png")));
			}
		});

		btn_login.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/sign-in-btn.png")));
		p_outputEmailLogin.add(btn_login);

		lblimapGif = new JLabel("");
		lblimapGif.setBounds(102, 166, 48, 41);
		p_outputEmailLogin.add(lblimapGif);
		lblimapGif.setVisible(false);
		lblimapGif.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/loading.gif")));

		passwordField = new JPasswordField();
		passwordField.setBounds(20, 86, 217, 20);
		p_outputEmailLogin.add(passwordField);

		lblNewLabel_Password = new JLabel("Password");
		lblNewLabel_Password.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_Password.setBounds(84, 68, 71, 14);
		p_outputEmailLogin.add(lblNewLabel_Password);

		outputUsernameField = new JTextField();
		outputUsernameField.setBounds(20, 43, 217, 20);
		p_outputEmailLogin.add(outputUsernameField);
		outputUsernameField.setVisible(false);
		outputUsernameField.setColumns(10);

		lblUsername_1 = new JLabel("");
		lblUsername_1.setBounds(44, 7, 32, 32);
		lblUsername_1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/username.png")));
		lblUsername_1.setVisible(false);
		p_outputEmailLogin.add(lblUsername_1);

		txtCloudhostgatorcom = new JTextField();
		txtCloudhostgatorcom.setBounds(21, 135, 136, 20);
		txtCloudhostgatorcom.setVisible(false);
		p_outputEmailLogin.add(txtCloudhostgatorcom);
		txtCloudhostgatorcom.setColumns(10);

		lbl_Hostoutput = new JLabel("Host");
		lbl_Hostoutput.setVisible(false);
		lbl_Hostoutput.setFont(new Font("Tahoma", Font.BOLD, 11));
		lbl_Hostoutput.setBounds(25, 117, 46, 14);
		p_outputEmailLogin.add(lbl_Hostoutput);

		textField_portOutput = new JTextField();
		textField_portOutput.setBounds(167, 135, 86, 20);
		textField_portOutput.setVisible(false);
		p_outputEmailLogin.add(textField_portOutput);
		textField_portOutput.setColumns(10);

		lblNewLabel = new JLabel("Port");
		lblNewLabel.setVisible(false);
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel.setBounds(167, 117, 46, 14);
		p_outputEmailLogin.add(lblNewLabel);

		r_rtf = new JRadioButton("RTF");
		buttonGroup.add(r_rtf);
		r_rtf.setBackground(Color.WHITE);
		r_rtf.setBounds(449, 7, 56, 23);
		r_rtf.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fileSavingFormatEvent();
				outputSource = OutputSource.RTF.name();
			}
		});

		p_outputSavingformat.add(r_rtf);

		r_xps = new JRadioButton("XPS");
		r_xps.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				fileSavingFormatEvent();
				outputSource = OutputSource.XPS.name();
			}
		});

		buttonGroup.add(r_xps);
		r_xps.setBackground(Color.WHITE);
		r_xps.setBounds(6, 33, 56, 23);
		p_outputSavingformat.add(r_xps);

		r_emf = new JRadioButton("EMF");
		buttonGroup.add(r_emf);
		r_emf.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.EMF.name();
			}
		});
		r_emf.setBackground(Color.WHITE);
		r_emf.setBounds(68, 33, 47, 23);
		p_outputSavingformat.add(r_emf);

		r_docx = new JRadioButton("DOCX");
		buttonGroup.add(r_docx);
		r_docx.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.DOCX.name();
			}
		});
		r_docx.setBackground(Color.WHITE);
		r_docx.setBounds(117, 33, 56, 23);
		p_outputSavingformat.add(r_docx);

		r_jpeg = new JRadioButton("JPEG");
		buttonGroup.add(r_jpeg);
		r_jpeg.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.JPEG.name();
			}
		});
		r_jpeg.setBackground(Color.WHITE);
		r_jpeg.setBounds(175, 33, 56, 23);
		p_outputSavingformat.add(r_jpeg);

		r_docm = new JRadioButton("DOCM");
		buttonGroup.add(r_docm);
		r_docm.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.DOCM.name();
			}
		});
		r_docm.setBackground(Color.WHITE);
		r_docm.setBounds(231, 33, 56, 23);
		p_outputSavingformat.add(r_docm);

		r_text = new JRadioButton("TEXT");
		buttonGroup.add(r_text);
		r_text.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.TEXT.name();
			}
		});
		r_text.setBackground(Color.WHITE);
		r_text.setBounds(296, 33, 56, 23);
		p_outputSavingformat.add(r_text);

		r_tiff = new JRadioButton("TIFF");
		buttonGroup.add(r_tiff);
		r_tiff.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.TIFF.name();
			}
		});
		r_tiff.setBackground(Color.WHITE);
		r_tiff.setBounds(368, 33, 56, 23);
		p_outputSavingformat.add(r_tiff);

		r_png = new JRadioButton("PNG");
		buttonGroup.add(r_png);
		r_png.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.PNG.name();
			}
		});
		r_png.setBackground(Color.WHITE);
		r_png.setBounds(449, 33, 56, 23);
		p_outputSavingformat.add(r_png);

		r_svg = new JRadioButton("SVG");
		buttonGroup.add(r_svg);
		r_svg.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.SVG.name();
			}
		});
		r_svg.setBackground(Color.WHITE);
		r_svg.setBounds(6, 59, 56, 23);
		p_outputSavingformat.add(r_svg);

		r_epub = new JRadioButton("EPUB");
		buttonGroup.add(r_epub);
		r_epub.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.EPUB.name();
			}
		});
		r_epub.setBackground(Color.WHITE);
		r_epub.setBounds(67, 59, 55, 23);
		p_outputSavingformat.add(r_epub);

		r_dotm = new JRadioButton("DOTM");
		buttonGroup.add(r_dotm);
		r_dotm.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.DOTM.name();
			}
		});
		r_dotm.setBackground(Color.WHITE);
		r_dotm.setBounds(120, 59, 56, 23);
		p_outputSavingformat.add(r_dotm);

		r_ott = new JRadioButton("OTT");
		buttonGroup.add(r_ott);
		r_ott.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.OTT.name();
			}
		});
		r_ott.setBackground(Color.WHITE);
		r_ott.setBounds(296, 59, 56, 23);
		p_outputSavingformat.add(r_ott);

		r_gif = new JRadioButton("GIF");
		buttonGroup.add(r_gif);
		r_gif.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.GIF.name();
			}
		});
		r_gif.setBackground(Color.WHITE);
		r_gif.setBounds(231, 59, 47, 23);
		p_outputSavingformat.add(r_gif);

		r_bmp = new JRadioButton("BMP");
		buttonGroup.add(r_bmp);
		r_bmp.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.BMP.name();
			}
		});
		r_bmp.setBackground(Color.WHITE);
		r_bmp.setBounds(175, 59, 47, 23);
		p_outputSavingformat.add(r_bmp);

		r_wordml = new JRadioButton("WORLD ML");
		buttonGroup.add(r_wordml);
		r_wordml.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.WORLD_ML.name();
			}
		});
		r_wordml.setBackground(Color.WHITE);
		r_wordml.setBounds(368, 59, 80, 23);
		p_outputSavingformat.add(r_wordml);

		r_odt = new JRadioButton("ODT");
		buttonGroup.add(r_odt);
		r_odt.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.ODT.name();
			}
		});
		r_odt.setBackground(Color.WHITE);
		r_odt.setBounds(449, 59, 56, 23);
		p_outputSavingformat.add(r_odt);

		r_csv = new JRadioButton("CSV");
		buttonGroup.add(r_csv);
		r_csv.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				fileSavingFormatEvent();
				outputSource = OutputSource.CSV.name();
			}
		});
		r_csv.setBackground(Color.WHITE);
		r_csv.setBounds(6, 85, 56, 23);
		p_outputSavingformat.add(r_csv);

		r_imap = new JRadioButton("Imap");
		buttonGroup.add(r_imap);
		r_imap.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.IMAP.name();
			}
		});
		r_imap.setBackground(Color.WHITE);
		r_imap.setBounds(96, 238, 63, 23);
		p_outputSavingformat.add(r_imap);

		r_hostgator = new JRadioButton("Hosgator");
		buttonGroup.add(r_hostgator);
		r_hostgator.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				radioButtionEvent(e);
				outputSource = OutputSource.HostGator.name();
			}
		});

		r_hostgator.setBackground(Color.WHITE);
		r_hostgator.setBounds(187, 160, 74, 23);
		p_outputSavingformat.add(r_hostgator);

		r_hotmail = new JRadioButton("Hotmail");
		r_hotmail.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.Hotmail.name();
			}
		});
		buttonGroup.add(r_hotmail);
		r_hotmail.setBackground(Color.WHITE);
		r_hotmail.setBounds(190, 186, 99, 23);
		p_outputSavingformat.add(r_hotmail);

		r_aws = new JRadioButton("Aws");
		buttonGroup.add(r_aws);

		r_aws.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.AWS.name();
			}
		});

		r_aws.setBackground(Color.WHITE);
		r_aws.setBounds(187, 212, 47, 23);
		p_outputSavingformat.add(r_aws);

		r_icloud = new JRadioButton("Icloud");
		buttonGroup.add(r_icloud);
		r_icloud.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.ICLOUD.name();
			}
		});
		r_icloud.setBackground(Color.WHITE);
		r_icloud.setBounds(7, 239, 56, 23);
		p_outputSavingformat.add(r_icloud);

		r_gmail_app = new JRadioButton("Gmail App");
		buttonGroup.add(r_gmail_app);
		r_gmail_app.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
				outputSource = OutputSource.GMAIL_APP.name();
			}
		});
		r_gmail_app.setBackground(Color.WHITE);
		r_gmail_app.setBounds(187, 238, 107, 23);
		p_outputSavingformat.add(r_gmail_app);
		lblNewLabel_Password.setVisible(false);
		passwordField.setVisible(false);
		btn_login.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				LogUtils.setTextToLogScreen(textPane_log, logger, "Selected Output Source  : " + outputSource);
				output_password = new String(passwordField.getPassword()).trim();
				output_userName = new String(outputUsernameField.getText()).trim();
				output_imapHost = new String(txtCloudhostgatorcom.getText()).trim();
				output_portNo = new String(textField_portOutput.getText()).trim();
				boolean check = true;
				LogUtils.setTextToLogScreen(textPane_log, logger, "Connecting with : " + output_userName);
				if (r_imap.isSelected() || r_hostgator.isSelected()) {
					if (output_imapHost.isEmpty() && output_portNo.isEmpty()) {
						check = false;
					}
				}

				if (!output_password.isEmpty() && !output_userName.isEmpty() && check
						&& output_userName.contains("@")) {

					Thread threadTable = new Thread(new Runnable() {

						@Override
						public void run() {
							btnDisable();
							try {

								if (r_gmail_app.isSelected()) {
									//passwordField.setText("");
						
									OutputGmailAppLogin();									
									getoutputGmailAppService();
									outputGmailAppInitialfolderCreation();

								} else {
									if (clientforimap_Output != null && iconnforimap_input != null) {
										clientforimap_Output.setUseMultiConnection(MultiConnectionMode.Disable);
										clientforimap_Output.close();
										clientforimap_Output.dispose();
										System.out.println("connection is closed");
									}
									clientforimap_Output = outputImapConnection();
									LogUtils.setTextToLogScreen(textPane_log, logger,
											"Connecting done with : " + output_userName);
									outputImapInitialfolderCreation(clientforimap_Output);

								}
								CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
								card.show(EmailWizardApplication.CardLayout, "GoogleDownloading_4");
								btnEnabled();

							} catch (Exception ex) {
								ex.printStackTrace();
								btnEnabled();
								ExceptionHandler exceptionHandler = new ExceptionHandler(ex, frame);
								exceptionHandler.loginExceptionHandler();
							}

						}

					});
					threadTable.start();

				} else {

					LogUtils.setTextToLogScreen(textPane_log, logger, "Error:fields cannot be empty!");

					JOptionPane.showMessageDialog(EmailWizardApplication.this,
							"Fields cannot be empty or please check your entered details are correct!",
							ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

				}

			}
		});
		btn_login.setVisible(false);
		lblNewLabel_Useranme.setVisible(false);

		DateFilterPanel = new JPanel();
		DateFilterPanel.setBackground(Color.WHITE);
		SavingOptionPanel_1.add(DateFilterPanel, "panel_dateFilter");
		DateFilterPanel.setLayout(null);

		chckbxSkipDuplicate = new JCheckBox("Skip Duplicate Email(s)");
		chckbxSkipDuplicate.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == ItemEvent.SELECTED) {

					chckbxSkip_subject.setEnabled(true);
					chckbxSkip_date.setEnabled(true);
					chckbxSkip_from.setEnabled(true);
					chckbxSkip_body.setEnabled(true);

				} else if (e.getStateChange() == ItemEvent.DESELECTED) {

					chckbxSkip_subject.setEnabled(false);
					chckbxSkip_date.setEnabled(false);
					chckbxSkip_from.setEnabled(false);
					chckbxSkip_body.setEnabled(false);

				}
			}
		});
		chckbxSkipDuplicate.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxSkipDuplicate.setBounds(6, 4, 168, 23);
		DateFilterPanel.add(chckbxSkipDuplicate);
		chckbxSkipDuplicate.setBackground(Color.WHITE);

		chckbxNamingconvention = new JCheckBox("Naming Convention");
		chckbxNamingconvention.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxNamingconvention.setBounds(6, 61, 142, 23);
		DateFilterPanel.add(chckbxNamingconvention);
		chckbxNamingconvention.setBackground(Color.WHITE);

		comboBoxNamingConvention = new JComboBox();
		comboBoxNamingConvention.setBounds(154, 63, 248, 20);
		DateFilterPanel.add(comboBoxNamingConvention);
		comboBoxNamingConvention.addItem("Subject");
		comboBoxNamingConvention.addItem("Subject_Date(DD-MM-YYYY)");
		comboBoxNamingConvention.addItem("Subject_Date(MM-DD-YYYY)");
		comboBoxNamingConvention.addItem("Subject_Date(YYYY-MM-DD)");
		comboBoxNamingConvention.addItem("Subject_Date(YYYY-DD-MM)");
		comboBoxNamingConvention.addItem("(DD-MM-YYYY)Date_Subject");
		comboBoxNamingConvention.addItem("(MM-DD-YYYY)Date_Subject");
		comboBoxNamingConvention.addItem("(YYYY-MM-DD)Date_Subject");
		comboBoxNamingConvention.addItem("(YYYY-DD-MM)Date_Subject");
		comboBoxNamingConvention.addItem("From_Subject_Date(DD-MM-YYYY)");
		comboBoxNamingConvention.addItem("From_Subject_Date(MM-DD-YYYY)");
		comboBoxNamingConvention.addItem("From_Subject_Date(YYYY-MM-DD)");
		comboBoxNamingConvention.addItem("From_Subject_Date(YYYY-DD-MM)");
		comboBoxNamingConvention.addItem("(DD-MM-YYYY)Date_From_Subject");
		comboBoxNamingConvention.addItem("(MM-DD-YYYY)Date_From_Subject");
		comboBoxNamingConvention.addItem("(YYYY-MM-DD)Date_From_Subject");
		comboBoxNamingConvention.addItem("(YYYY-DD-MM)Date_From_Subject");

		panel_5 = new JPanel();
		panel_5.setBounds(38, 193, 344, 57);
		DateFilterPanel.add(panel_5);
		panel_5.setEnabled(false);
		panel_5.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "", TitledBorder.LEFT,
				TitledBorder.TOP, null, new Color(0, 0, 255)));
		panel_5.setBackground(Color.WHITE);
		panel_5.setLayout(null);

		start_dateChooser = new JDateChooser();
		start_dateChooser.setEnabled(false);
		JTextFieldDateEditor editorStart = (JTextFieldDateEditor) start_dateChooser.getDateEditor();
		start_dateChooser.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar start_dateCalendar = Calendar.getInstance();
				Date startdate = start_dateCalendar.getTime();
				start_dateChooser.setMaxSelectableDate(startdate);

			}
		});
		start_dateChooser.setBounds(15, 28, 146, 20);
		panel_5.add(start_dateChooser);

		end_dateChooser = new JDateChooser();
		JTextFieldDateEditor editor = (JTextFieldDateEditor) end_dateChooser.getDateEditor();
		end_dateChooser.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				Calendar end_dateCalendar = Calendar.getInstance();
				Date enddate = end_dateCalendar.getTime();
				end_dateChooser.setMaxSelectableDate(enddate);

			}
		});
		end_dateChooser.setBounds(190, 28, 146, 20);
		end_dateChooser.setEnabled(false);
		panel_5.add(end_dateChooser);

		JLabel lblStartDate = new JLabel("Start Date");
		lblStartDate.setFont(new Font("Tahoma", Font.BOLD, 10));
		lblStartDate.setBounds(49, 9, 56, 14);
		panel_5.add(lblStartDate);

		JLabel lblEndDate = new JLabel("End Date");
		lblEndDate.setFont(new Font("Tahoma", Font.BOLD, 10));
		lblEndDate.setBounds(214, 9, 56, 14);
		panel_5.add(lblEndDate);

		rdbtnDateFilter = new JRadioButton("Date Filter");
		rdbtnDateFilter.setBounds(6, 163, 88, 23);
		DateFilterPanel.add(rdbtnDateFilter);
		rdbtnDateFilter.setFont(new Font("Tahoma", Font.BOLD, 11));
		rdbtnDateFilter.setForeground(Color.BLACK);
		rdbtnDateFilter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				if (rdbtnDateFilter.isSelected()) {
					panel_5.setEnabled(true);
					start_dateChooser.setEnabled(true);
					end_dateChooser.setEnabled(true);

				} else {
					panel_5.setEnabled(false);
					start_dateChooser.setEnabled(false);
					end_dateChooser.setEnabled(false);
				}

			}
		});
		rdbtnDateFilter.setBackground(Color.WHITE);

		checkBoxSplitPst = new JCheckBox("Split Pst ");
		checkBoxSplitPst.setFont(new Font("Tahoma", Font.BOLD, 11));
		checkBoxSplitPst.setBounds(6, 93, 88, 23);
		DateFilterPanel.add(checkBoxSplitPst);
		checkBoxSplitPst.setVisible(false);
		checkBoxSplitPst.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				if (checkBoxSplitPst.isSelected()) {
					radioButtonMB.setVisible(true);
					rdbtnGb.setVisible(true);
					spinner_GB.setVisible(true);
					spinner_MB.setVisible(true);
				} else {
					radioButtonMB.setVisible(false);
					rdbtnGb.setVisible(false);
					spinner_GB.setVisible(false);
					spinner_MB.setVisible(false);
				}

			}
		});
		checkBoxSplitPst.setBackground(Color.WHITE);

		radioButtonMB = new JRadioButton("MB");
		radioButtonMB.setBounds(161, 92, 41, 23);
		DateFilterPanel.add(radioButtonMB);
		radioButtonMB.setSelected(true);
		radioButtonMB.setVisible(false);
		buttonGroup_1.add(radioButtonMB);
		radioButtonMB.setBackground(Color.WHITE);

		spinner_MB = new JSpinner();
		spinner_MB.setBounds(208, 92, 47, 20);
		DateFilterPanel.add(spinner_MB);
		spinner_MB.setVisible(false);

		spinner_MB.setModel(new SpinnerNumberModel(new Integer(1), new Integer(0), null, new Integer(1)));

		spinner_GB = new JSpinner();
		spinner_GB.setBounds(317, 92, 49, 20);
		DateFilterPanel.add(spinner_GB);
		spinner_GB.setVisible(false);
		spinner_GB.setModel(new SpinnerNumberModel(new Integer(1), new Integer(1), null, new Integer(1)));

		rdbtnGb = new JRadioButton("GB");
		rdbtnGb.setBounds(269, 91, 41, 23);
		DateFilterPanel.add(rdbtnGb);
		rdbtnGb.setVisible(false);
		buttonGroup_1.add(rdbtnGb);
		rdbtnGb.setBackground(Color.WHITE);

		chckbxSaveSeperateAttachments = new JCheckBox("Save Attachments In Seperate Folder");
		chckbxSaveSeperateAttachments.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxSaveSeperateAttachments.setBackground(Color.WHITE);
		chckbxSaveSeperateAttachments.setBounds(6, 126, 286, 23);
		DateFilterPanel.add(chckbxSaveSeperateAttachments);

		chckbxSkip_body = new JCheckBox("Body");
		chckbxSkip_body.setEnabled(false);
		chckbxSkip_body.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSkip_body.setBackground(Color.WHITE);
		chckbxSkip_body.setBounds(48, 29, 59, 23);
		DateFilterPanel.add(chckbxSkip_body);

		chckbxSkip_subject = new JCheckBox("Subject");
		chckbxSkip_subject.setEnabled(false);
		chckbxSkip_subject.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSkip_subject.setBackground(Color.WHITE);
		chckbxSkip_subject.setBounds(109, 29, 70, 23);
		DateFilterPanel.add(chckbxSkip_subject);

		chckbxSkip_from = new JCheckBox("From");
		chckbxSkip_from.setEnabled(false);
		chckbxSkip_from.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSkip_from.setBackground(Color.WHITE);
		chckbxSkip_from.setBounds(181, 29, 59, 23);
		DateFilterPanel.add(chckbxSkip_from);

		chckbxSkip_date = new JCheckBox("Date & Time");
		chckbxSkip_date.setEnabled(false);
		chckbxSkip_date.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSkip_date.setBackground(Color.WHITE);
		chckbxSkip_date.setBounds(242, 28, 124, 23);
		DateFilterPanel.add(chckbxSkip_date);
		JFormattedTextField txt_GB = ((JSpinner.NumberEditor) spinner_GB.getEditor()).getTextField();
		JFormattedTextField txt_MB = ((JSpinner.NumberEditor) spinner_MB.getEditor()).getTextField();
		comboBoxNamingConvention.setVisible(true);
		r_yandex.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {

				radioButtionEvent(e);
			}
		});
		((NumberFormatter) txt_MB.getFormatter()).setAllowsInvalid(false);
		((NumberFormatter) txt_MB.getFormatter()).setMinimum(1);
		((NumberFormatter) txt_GB.getFormatter()).setAllowsInvalid(false);
		((NumberFormatter) txt_GB.getFormatter()).setMinimum(1);

		JButton b_filterOptions = new JButton("");
		b_filterOptions.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				CardLayout card = (CardLayout) SavingOptionPanel_1.getLayout();
				card.show(SavingOptionPanel_1, "panel_dateFilter");
			}
		});
		b_filterOptions.setBounds(661, 13, 117, 36);

		b_filterOptions.setRolloverEnabled(false);
		b_filterOptions.setRequestFocusEnabled(false);
		b_filterOptions.setOpaque(false);
		b_filterOptions.setFocusable(false);
		b_filterOptions.setFocusTraversalKeysEnabled(false);
		b_filterOptions.setFocusPainted(false);
		b_filterOptions.setDefaultCapable(false);
		b_filterOptions.setContentAreaFilled(false);
		b_filterOptions.setBorderPainted(false);
		b_filterOptions.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				b_filterOptions.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/data-filter.png")));
			}

			public void mouseExited(MouseEvent e) {

				b_filterOptions
						.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/data-filter-hvr.png")));
			}
		});

		b_filterOptions.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/data-filter.png")));

		SavingOptionPanel.add(b_filterOptions);

		c_email = new JCheckBox("Email");
		c_email.setBounds(0, 43, 58, 23);
		SavingOptionPanel.add(c_email);
		c_email.setSelected(true);
		c_email.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				if (c_email.isSelected()) {
					SavingOptionPanel_1.setVisible(true);

				} else {
					SavingOptionPanel_1.setVisible(false);
				}
			}
		});
		c_email.setFont(new Font("Tahoma", Font.BOLD, 11));
		c_email.setBackground(Color.WHITE);

		l_gmail = new JLabel("");
		l_gmail.setBounds(0, 3, 49, 36);
		SavingOptionPanel.add(l_gmail);
		l_gmail.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/gmail.png")));

		l_contact = new JLabel("");
		l_contact.setBounds(104, 3, 46, 42);
		SavingOptionPanel.add(l_contact);
		l_contact.setVisible(false);
		l_contact.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/contact.png")));

		c_contact = new JCheckBox("Contact");
		c_contact.setBounds(101, 45, 75, 23);
		SavingOptionPanel.add(c_contact);
		c_contact.setVisible(false);
		c_contact.setFont(new Font("Tahoma", Font.BOLD, 11));
		c_contact.setBackground(Color.WHITE);

		l_calendar = new JLabel("");
		l_calendar.setBounds(207, -2, 49, 49);
		SavingOptionPanel.add(l_calendar);
		l_calendar.setVisible(false);
		l_calendar.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/calender.png")));

		c_calendar = new JCheckBox("Calendar");
		c_calendar.setBounds(194, 45, 83, 23);
		SavingOptionPanel.add(c_calendar);
		c_calendar.setVisible(false);
		c_calendar.setFont(new Font("Tahoma", Font.BOLD, 11));
		c_calendar.setBackground(Color.WHITE);

		l_drive = new JLabel("");
		l_drive.setBounds(315, 1, 49, 49);
		SavingOptionPanel.add(l_drive);
		l_drive.setVisible(false);
		l_drive.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/gdrive.png")));

		c_drive = new JCheckBox("Drive");
		c_drive.setBounds(314, 48, 63, 16);
		SavingOptionPanel.add(c_drive);
		c_drive.setVisible(false);
		c_drive.setFont(new Font("Tahoma", Font.BOLD, 11));
		c_drive.setBackground(Color.WHITE);

		l_photos = new JLabel("");
		l_photos.setBounds(425, 4, 49, 47);
		SavingOptionPanel.add(l_photos);
		l_photos.setVisible(false);
		l_photos.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/photo.png")));

		c_photos = new JCheckBox("Photos");
		c_photos.setBounds(420, 53, 70, 14);
		SavingOptionPanel.add(c_photos);
		c_photos.setFont(new Font("Tahoma", Font.BOLD, 11));
		c_photos.setVisible(false);
		c_photos.setBackground(Color.WHITE);

		JButton b_backup = new JButton();
		b_backup.setBounds(547, 13, 113, 39);
		b_backup.setRolloverEnabled(false);
		b_backup.setRequestFocusEnabled(false);
		b_backup.setOpaque(false);
		b_backup.setFocusable(false);
		b_backup.setFocusTraversalKeysEnabled(false);
		b_backup.setFocusPainted(false);
		b_backup.setDefaultCapable(false);
		b_backup.setContentAreaFilled(false);
		b_backup.setBorderPainted(false);
		b_backup.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				b_backup.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/backup-option.png")));
			}

			public void mouseExited(MouseEvent e) {

				b_backup.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/backup-option-hvr.png")));
			}
		});

		b_backup.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/backup-option.png")));

		SavingOptionPanel.add(b_backup);
		b_backup.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				CardLayout card = (CardLayout) EmailWizardApplication.SavingOptionPanel_1.getLayout();
				card.show(EmailWizardApplication.SavingOptionPanel_1, "p_outputSavingformat");
			}
		});

		label_8 = new JLabel("");
		label_8.setBounds(0, 354, 780, 45);
		label_8.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/bottomn.png")));
		SavingFormatPanel_P4.add(label_8);

		JPanel MigrationPanel_Table_P5 = new JPanel();
		MigrationPanel_Table_P5.setBackground(Color.WHITE);
		CardLayout.add(MigrationPanel_Table_P5, "GoogleDownloading_4");
		MigrationPanel_Table_P5.setLayout(null);

		btnDownloading = new JButton("");
		btnDownloading.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				boolean ischeckEmailClientSelected = OutputSource.imapClientOutputFormat.contains(emailClientSelectedFormatAtOutput());
						

				if (!textField_DownloadingPath.getText().isEmpty() || ischeckEmailClientSelected) {

					Thread threadTable = new Thread(new Runnable() {

						@Override
						public void run() {

							LogUtils.setTextToLogScreen(textPane_log, logger, "Downlaoding Process Started");
							if (selectedInput.equals(InputSource.GSUITE.getValue())) {
								GSuiteMigration();
							} else if (selectedInput.equals(InputSource.MS_Office_365.getValue())) {
								// setFoldeList(MsOfficeMigrationTree(folderIdlist));
								MsOfficeMigration();
							} else if (selectedInput.equals(InputSource.GMAIL_APP.getValue())) {
								GmailAppMigration();
							} else if (selectedInput.equals(InputSource.Bulk.getValue())) {
								ImapBulkMigration();
							} else {
								ImapMigration();

							}

						}
					});
					threadTable.start();

				} else {

					System.out.println("Please select Path!!!");
					JOptionPane.showMessageDialog(EmailWizardApplication.this, "Please select Path!!",
							ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
				}
			}
		});
		btnDownloading.setBounds(666, 363, 114, 34);

		btnDownloading.setRolloverEnabled(false);
		btnDownloading.setRequestFocusEnabled(false);
		btnDownloading.setOpaque(false);
		btnDownloading.setFocusable(false);
		btnDownloading.setFocusTraversalKeysEnabled(false);
		btnDownloading.setFocusPainted(false);
		btnDownloading.setDefaultCapable(false);
		btnDownloading.setContentAreaFilled(false);
		btnDownloading.setBorderPainted(false);
		btnDownloading.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnDownloading
						.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/download-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnDownloading.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/download-btn.png")));
			}
		});

		btnDownloading.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/download-btn.png")));

		MigrationPanel_Table_P5.add(btnDownloading);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setBounds(0, 0, 780, 208);
		MigrationPanel_Table_P5.add(scrollPane_1);

		table_Downloading = new JTable() {
			/**
			 *
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};
		table_Downloading.getTableHeader().setReorderingAllowed(false);
		table_Downloading.setBackground(Color.WHITE);
		table_Downloading.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "<HTML><B>Folders</B></HTML>", "<HTML><B>Folder Name/Count</B></HTML>",
						"<HTML><B>Error Count</B></HTML>", "<HTML><B>Mail Count</B></HTML>",
						"<HTML><B>Total Mail In Folder</B></HTML>" }));
		scrollPane_1.setViewportView(table_Downloading);

		btnBack_p4 = new JButton("");
		btnBack_p4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				DefaultTableModel dm = (DefaultTableModel) table_Downloading.getModel();
				while (dm.getRowCount() > 0) {
					dm.removeRow(0);
				}
				rownCount = 0;

				CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
				card.show(EmailWizardApplication.CardLayout, "GoogleDownloadOptions_3");
				btn_login.setEnabled(true);
			}
		});
		btnBack_p4.setBounds(0, 357, 124, 34);
		btnBack_p4.setRolloverEnabled(false);
		btnBack_p4.setRequestFocusEnabled(false);
		btnBack_p4.setOpaque(false);
		btnBack_p4.setFocusable(false);
		btnBack_p4.setFocusTraversalKeysEnabled(false);
		btnBack_p4.setFocusPainted(false);
		btnBack_p4.setDefaultCapable(false);
		btnBack_p4.setContentAreaFilled(false);
		btnBack_p4.setBorderPainted(false);
		btnBack_p4.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnBack_p4.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnBack_p4.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
			}
		});

		btnBack_p4.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
		MigrationPanel_Table_P5.add(btnBack_p4);

		JPanel SavePathPanel = new JPanel();
		SavePathPanel.setBackground(Color.WHITE);
		SavePathPanel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		SavePathPanel.setBounds(10, 216, 760, 136);
		MigrationPanel_Table_P5.add(SavePathPanel);
		SavePathPanel.setLayout(null);

		btnDownloadingPath = new JButton("");
		btnDownloadingPath.setBounds(642, 97, 108, 28);
		SavePathPanel.add(btnDownloadingPath);
		btnDownloadingPath.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				JFileChooser jFileChooser = new JFileChooser();

				jFileChooser.setBackground(Color.WHITE);

				jFileChooser.setAcceptAllFileFilterUsed(false);

				jFileChooser.setMultiSelectionEnabled(true);

				jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				jFileChooser.showOpenDialog(EmailWizardApplication.this);

				File file = jFileChooser.getSelectedFile();
				if (!(file == null)) {

					String destination = file.getAbsolutePath();
					textField_DownloadingPath.setText(destination);

				}

			}
		});

		btnDownloadingPath.setRolloverEnabled(false);
		btnDownloadingPath.setRequestFocusEnabled(false);
		btnDownloadingPath.setOpaque(false);
		btnDownloadingPath.setFocusable(false);
		btnDownloadingPath.setFocusTraversalKeysEnabled(false);
		btnDownloadingPath.setFocusPainted(false);
		btnDownloadingPath.setDefaultCapable(false);
		btnDownloadingPath.setContentAreaFilled(false);
		btnDownloadingPath.setBorderPainted(false);
		btnDownloadingPath.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnDownloadingPath
						.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/dest-path-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnDownloadingPath
						.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/dest-path-btn.png")));
			}
		});

		btnDownloadingPath.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/dest-path-btn.png")));

		progressBar_Downloading = new JProgressBar();
		progressBar_Downloading.setBounds(12, 61, 622, 28);
		SavePathPanel.add(progressBar_Downloading);

		lblDownloading = new JLabel("Downloading Status");
		lblDownloading.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblDownloading.setBounds(10, 6, 131, 14);
		SavePathPanel.add(lblDownloading);
		lblDownloading.setVisible(false);
		lblDownloading.setForeground(Color.BLACK);

		textField_DownloadingPath = new JTextField();
		textField_DownloadingPath.setEditable(false);
		textField_DownloadingPath.setBounds(11, 97, 622, 28);
		SavePathPanel.add(textField_DownloadingPath);
		textField_DownloadingPath.setColumns(10);

		ProgressBarPanel = new JPanel();
		ProgressBarPanel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		ProgressBarPanel.setBackground(Color.WHITE);
		ProgressBarPanel.setBounds(10, 26, 740, 28);
		SavePathPanel.add(ProgressBarPanel);
		ProgressBarPanel.setLayout(null);

		dataName = new JLabel("");
		dataName.setBounds(10, 7, 106, 14);
		ProgressBarPanel.add(dataName);
		dataName.setFont(new Font("Tahoma", Font.BOLD, 12));

		downloadingFileName = new JLabel("");
		downloadingFileName.setBounds(148, 7, 592, 14);
		ProgressBarPanel.add(downloadingFileName);

		btnStop = new JButton("");
		btnStop.setBounds(638, 56, 114, 34);
		SavePathPanel.add(btnStop);
		btnStop.setRolloverEnabled(false);
		btnStop.setRequestFocusEnabled(false);
		btnStop.setOpaque(false);
		btnStop.setFocusable(false);
		btnStop.setFocusTraversalKeysEnabled(false);
		btnStop.setFocusPainted(false);
		btnStop.setDefaultCapable(false);
		btnStop.setContentAreaFilled(false);
		btnStop.setBorderPainted(false);
		btnStop.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnStop.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/stop-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnStop.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/stop-btn.png")));
			}
		});

		btnStop.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/stop-btn.png")));

		lblNoInternetConnection = new JLabel("No Internet Connection....Please check your internet  connection.");
		lblNoInternetConnection.setBounds(201, 6, 331, 14);
		SavePathPanel.add(lblNoInternetConnection);
		lblNoInternetConnection.setVisible(false);
		lblNoInternetConnection.setForeground(Color.RED);

		JButton btnLogScreen = new JButton("Log Screen");
		btnLogScreen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
				card.show(EmailWizardApplication.CardLayout, "panel_LogScreen");
			}
		});
		btnLogScreen.setBounds(352, 359, 89, 34);
		MigrationPanel_Table_P5.add(btnLogScreen);

		label_9 = new JLabel("");
		label_9.setBounds(0, 357, 780, 40);
		label_9.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/bottomn.png")));
		MigrationPanel_Table_P5.add(label_9);

		MsOfficePanel_P6 = new JPanel();
		MsOfficePanel_P6.setBackground(Color.WHITE);
		CardLayout.add(MsOfficePanel_P6, "p_msoffice");
		MsOfficePanel_P6.setLayout(null);

		r_mailbox = new JRadioButton("Mailbox Folder");
		r_mailbox.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				c_DeepFolderTraversal.setVisible(true);
			}
		});
		r_mailbox.setBounds(124, 141, 135, 23);
		r_mailbox.setSelected(true);
		buttonGroup_3.add(r_mailbox);
		r_mailbox.setBackground(Color.WHITE);
		r_mailbox.setFont(new Font("Tahoma", Font.BOLD, 12));
		MsOfficePanel_P6.add(r_mailbox);

		r_public = new JRadioButton("Public Folder");
		r_public.setBounds(294, 141, 109, 23);
		r_public.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				c_DeepFolderTraversal.setSelected(false);
				c_DeepFolderTraversal.setVisible(false);
			}
		});
		buttonGroup_3.add(r_public);
		r_public.setBackground(Color.WHITE);
		r_public.setFont(new Font("Tahoma", Font.BOLD, 12));
		MsOfficePanel_P6.add(r_public);

		r_archive = new JRadioButton("Archive Folder");
		r_archive.setBounds(445, 141, 197, 23);
		r_archive.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				c_DeepFolderTraversal.setVisible(true);
			}
		});
		buttonGroup_3.add(r_archive);
		r_archive.setBackground(Color.WHITE);
		r_archive.setFont(new Font("Tahoma", Font.BOLD, 12));
		MsOfficePanel_P6.add(r_archive);

		b_next1 = new JButton("");
		b_next1.setBounds(661, 361, 109, 34);
		b_next1.setRolloverEnabled(false);
		b_next1.setRequestFocusEnabled(false);
		b_next1.setOpaque(false);
		b_next1.setFocusable(false);
		b_next1.setFocusTraversalKeysEnabled(false);
		b_next1.setFocusPainted(false);
		b_next1.setDefaultCapable(false);
		b_next1.setContentAreaFilled(false);
		b_next1.setBorderPainted(false);
		b_next1.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				b_next1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				b_next1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));
			}
		});

		b_next1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/next-btn.png")));

		b_next1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				DefaultTableModel model = (DefaultTableModel) EmailWizardApplication.table_UserDetails.getModel();

				DefaultTreeModel modeltree = (DefaultTreeModel) folderTree.getModel();
				DefaultMutableTreeNode root = new DefaultMutableTreeNode(input_userName);
				modeltree.setRoot(root);

				try {
					if (c_DeepFolderTraversal.isSelected()) {
//						ews.getFolder(service, model, true);
//						CardLayout card = (CardLayout) GoogleMaineFrame.CardLayout.getLayout();
//						card.show(GoogleMaineFrame.CardLayout, "GoogleUserDetailsPanel_2");

						ews.getFolderTree(service, root, model, false);
						CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
						card.show(EmailWizardApplication.CardLayout, "treePanel");

					} else {
						ews.getFolder(service, model, false);
						CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
						card.show(EmailWizardApplication.CardLayout, "GoogleUserDetailsPanel_2");
					}

				} catch (ServiceResponseException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					logger.warn("No Access to Folder!" + mailBox());

					JOptionPane.showMessageDialog(EmailWizardApplication.this,
							mailBox() + " could not be found in " + input_userName, ToolDetails.messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
				} catch (ServiceRequestException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					logger.warn(e1.getMessage());

					JOptionPane.showMessageDialog(EmailWizardApplication.this,
							mailBox() + " The request failed.The remote server returned an error: (401) try again",
							ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
				} catch (Exception e1) {
					e1.printStackTrace();
				}
			}
		});
		MsOfficePanel_P6.add(b_next1);

		b_back1 = new JButton("");
		b_back1.setBounds(3, 361, 124, 34);

		b_back1.setRolloverEnabled(false);
		b_back1.setRequestFocusEnabled(false);
		b_back1.setOpaque(false);
		b_back1.setFocusable(false);
		b_back1.setFocusTraversalKeysEnabled(false);
		b_back1.setFocusPainted(false);
		b_back1.setDefaultCapable(false);
		b_back1.setContentAreaFilled(false);
		b_back1.setBorderPainted(false);
		b_back1.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				b_back1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				b_back1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
			}
		});

		b_back1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));

		b_back1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
				card.show(EmailWizardApplication.CardLayout, "GoogleLoginPanel_1");
			}
		});
		MsOfficePanel_P6.add(b_back1);

		c_DeepFolderTraversal = new JCheckBox("Deep Folder Traversal ");
		c_DeepFolderTraversal.setFont(new Font("Tahoma", Font.BOLD, 11));
		c_DeepFolderTraversal.setBackground(Color.WHITE);
		c_DeepFolderTraversal.setBounds(263, 205, 170, 23);
		MsOfficePanel_P6.add(c_DeepFolderTraversal);

		JPanel MsOfficePanel_1 = new JPanel();
		MsOfficePanel_1.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		MsOfficePanel_1.setBackground(Color.WHITE);
		MsOfficePanel_1.setBounds(67, 88, 628, 110);
		MsOfficePanel_P6.add(MsOfficePanel_1);

		JLabel lblNewLabel_1 = new JLabel("");
		lblNewLabel_1.setBounds(0, 356, 780, 44);
		lblNewLabel_1.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/bottomn.png")));
		MsOfficePanel_P6.add(lblNewLabel_1);

		panel_LogScreen = new JPanel();
		CardLayout.add(panel_LogScreen, "panel_LogScreen");
		panel_LogScreen.setLayout(null);

		JScrollPane scrollPane_4 = new JScrollPane();
		scrollPane_4.setBounds(10, 11, 760, 346);
		panel_LogScreen.add(scrollPane_4);

		textPane_log = new JTextPane();

		// String someHtmlMessage = "<html><b style='color:blue;'>!!!!!!!!!---* Email
		// Wizard Application Started *----!!!!!!!!!!</b><html>";
		textPane_log.setText("!!!!!!!!!---*" + ToolDetails.messageboxtitle + " Started *----!!!!!!!!!!");
		scrollPane_4.setViewportView(textPane_log);
		btnNewButton = new JButton("");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
				card.show(EmailWizardApplication.CardLayout, "GoogleDownloading_4");
			}
		});
		btnNewButton.setBounds(303, 358, 169, 39);
		btnNewButton.setRolloverEnabled(false);
		btnNewButton.setRequestFocusEnabled(false);
		btnNewButton.setOpaque(false);
		btnNewButton.setFocusable(false);
		btnNewButton.setFocusTraversalKeysEnabled(false);
		btnNewButton.setFocusPainted(false);
		btnNewButton.setDefaultCapable(false);
		btnNewButton.setContentAreaFilled(false);
		btnNewButton.setBorderPainted(false);
		btnNewButton.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnNewButton.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnNewButton.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
			}
		});

		btnNewButton.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/back-btn.png")));
		panel_LogScreen.add(btnNewButton);

		JLabel lbllogbtm = new JLabel("");
		lbllogbtm.setBounds(0, 356, 780, 44);
		lbllogbtm.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/bottomn.png")));
		panel_LogScreen.add(lbllogbtm);
		btnStop.setVisible(false);
		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(EmailWizardApplication.this, warn, ToolDetails.messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(EmailWizardApplication.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					stop = true;
				}
			}

		});
		progressBar_Downloading.setVisible(false);

		JButton btnBuy = new JButton();
		btnBuy.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(ToolDetails.buyurl);
			}
		});
		btnBuy.setBounds(662, 5, 41, 32);
		btnBuy.setRolloverEnabled(false);
		btnBuy.setRequestFocusEnabled(false);
		btnBuy.setOpaque(false);
		btnBuy.setFocusable(false);
		btnBuy.setFocusTraversalKeysEnabled(false);
		btnBuy.setFocusPainted(false);
		btnBuy.setDefaultCapable(false);
		btnBuy.setContentAreaFilled(false);
		btnBuy.setBorderPainted(false);
		btnBuy.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnBuy.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/buy-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnBuy.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/buy-btn.png")));
			}
		});

		btnBuy.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/buy-btn.png")));
		contentPane.add(btnBuy);

		JButton btnInfo = new JButton("");
		btnInfo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				AboutDialog ab = null;
				if (demo) {
					ab = new AboutDialog(EmailWizardApplication.this, true, "Demo");
				} else {
					ab = new AboutDialog(EmailWizardApplication.this, true, "Full");

				}

				ab.setLocationRelativeTo(EmailWizardApplication.this);
				ab.setVisible(true);
			}
		});
		btnInfo.setBounds(739, -1, 41, 44);
		btnInfo.setRolloverEnabled(false);
		btnInfo.setRequestFocusEnabled(false);
		btnInfo.setOpaque(false);
		btnInfo.setFocusable(false);
		btnInfo.setFocusTraversalKeysEnabled(false);
		btnInfo.setFocusPainted(false);
		btnInfo.setDefaultCapable(false);
		btnInfo.setContentAreaFilled(false);
		btnInfo.setBorderPainted(false);
		btnInfo.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnInfo.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/info-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnInfo.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/info-btn.png")));
			}
		});

		btnInfo.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/info-btn.png")));
		contentPane.add(btnInfo);

		JButton btnAbout = new JButton();
		btnAbout.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(ToolDetails.helpurl);
			}
		});
		btnAbout.setBounds(701, 2, 41, 38);
		btnAbout.setRolloverEnabled(false);
		btnAbout.setRequestFocusEnabled(false);
		btnAbout.setOpaque(false);
		btnAbout.setFocusable(false);
		btnAbout.setFocusTraversalKeysEnabled(false);
		btnAbout.setFocusPainted(false);
		btnAbout.setDefaultCapable(false);
		btnAbout.setContentAreaFilled(false);
		btnAbout.setBorderPainted(false);
		btnAbout.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnAbout.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/about-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnAbout.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/about-btn.png")));
			}
		});

		btnAbout.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/about-btn.png")));
		contentPane.add(btnAbout);

		JButton btnActivation = new JButton("");
		if (demo) {
			btnActivation.setVisible(true);
		} else {
			btnActivation.setVisible(false);
		}
		btnActivation.setRolloverEnabled(false);
		btnActivation.setRequestFocusEnabled(false);
		btnActivation.setOpaque(false);
		btnActivation.setFocusable(false);
		btnActivation.setFocusTraversalKeysEnabled(false);
		btnActivation.setFocusPainted(false);
		btnActivation.setDefaultCapable(false);
		btnActivation.setContentAreaFilled(false);
		btnActivation.setBorderPainted(false);
		btnActivation.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {

				btnActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/key-act-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {

				btnActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/key-act-btn.png")));
			}
		});

		btnActivation.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/key-act-btn.png")));

		btnActivation.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				OnlineActivation af = new OnlineActivation(EmailWizardApplication.this, null, true);
				af.setLocationRelativeTo(null);
				af.setVisible(true);
				af.btnBack.setVisible(false);

			}
		});
		btnActivation.setBounds(620, 6, 41, 29);
		contentPane.add(btnActivation);

		l_topbar = new JLabel("");
		l_topbar.setBounds(0, 0, 780, 48);
		l_topbar.setIcon(new ImageIcon(EmailWizardApplication.class.getResource("/topbar.png")));

		contentPane.add(l_topbar);
	}

	public void buttonEnables() {

		progressBar_Downloading.setVisible(false);
		InputtxtUserName.setEnabled(true);
		txtServiceAccountIDorImapHostName.setEnabled(true);
		textField_p12FileAndPortNo.setEnabled(true);
		textField_DownloadingPath.setEnabled(true);
		btnBrowseP2File.setEnabled(true);
		btn_Login.setEnabled(true);
		btnBack_p2.setEnabled(true);
		btnNext_p2.setEnabled(true);
		btnBack_p3.setEnabled(true);
		btnNew_p3.setEnabled(true);
		chckbx_Proxy.setEnabled(true);
		btnDownloadingPath.setEnabled(true);
		textField_DownloadingPath.setEnabled(true);
		btnDownloading.setEnabled(true);
		btnBack_p4.setEnabled(true);
		label_LoginGif.setVisible(false);
		lblDownloading.setVisible(false);
		progressBar_Downloading.setVisible(false);
		btnStop.setVisible(false);
		EmailWizardApplication.downloadingFileName.setText("");
		dataName.setText("");
		ProgressBarPanel.setVisible(false);

	}

	public String downloadMSOfficeEmails(String folderName) throws Exception {

		if (r_mbox.isSelected()) {
			File destinationPath = new File(textField_DownloadingPath.getText() + File.separator + input_userName
					+ File.separator + outputSource + File.separator + mailBox() + File.separator
					+ FileNamingUtils.validFileNameForWindows(folderName));
			destinationPath.mkdirs();

			MboxrdStorageWriter mbox = new MboxrdStorageWriter(
					destinationPath.getAbsolutePath() + File.separator + destinationPath.getName() + ".mbx", false);

			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, destinationPath, mbox);

		} else if (EmailWizardApplication.r_pst.isSelected()) {

			pstfolderInfo = pst.getRootFolder().addSubFolder(folderName, true);
			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, pst, pstfolderInfo);
			return OutputSource.PST.name();

		} else if (r_office.isSelected() || r_aol.isSelected() || r_aws.isSelected() || r_gmail.isSelected()
				|| r_yahoo.isSelected() || r_yandex.isSelected() || r_zoho.isSelected() || r_icloud.isSelected()
				|| r_hotmail.isSelected()) {

			String parentPath = imapFolderPath + "/" + folderName;
			clientforimap_Output.createFolder(parentPath);

			if (r_aws.isSelected()) {
				clientforimap_Output.selectFolder(iconnforimap_Output, parentPath);
			} else {
				clientforimap_Output.selectFolder(iconnforimap_Output, parentPath);
				clientforimap_Output.subscribeFolder(iconnforimap_Output, parentPath);
			}

			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, parentPath,
					clientforimap_Output, iconnforimap_Output);

		}

		else if (r_imap.isSelected() || r_hostgator.isSelected()) {

			String parentPath = imapFolderPath + "." + folderName;

			clientforimap_Output.createFolder(iconnforimap_Output, "INBOX." + parentPath);
			clientforimap_Output.selectFolder(iconnforimap_Output, "INBOX." + parentPath);
			clientforimap_Output.subscribeFolder(iconnforimap_Output, "INBOX." + parentPath);

			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, parentPath,
					clientforimap_Output, iconnforimap_Output);

		}
		
		else if (r_gmail_app.isSelected()) {

			String parentPath = imapFolderPath + "/" + folderName;
			Label label = new Label();
			label.setName(parentPath);
			label.setLabelListVisibility("labelShow");
			label.setMessageListVisibility("show");
			String mainAccountUserName = new String(InputtxtUserName.getText()).trim();
			parentLabel = outputGmailService.users().labels().create(user, label).execute();
			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, parentPath, mainAccountUserName, parentLabel, outputGmailService);
					
		}

		else if (EmailWizardApplication.r_csv.isSelected()) {

			File destinationPath = new File(textField_DownloadingPath.getText() + File.separator + input_userName
					+ File.separator + outputSource + File.separator + mailBox() + File.separator
					+ FileNamingUtils.validFileNameForWindows(folderName));
			destinationPath.mkdirs();

			Workbook workbook = null;

			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, destinationPath,
					workbook);

		}

		else {
			File destinationPath = new File(textField_DownloadingPath.getText() + File.separator + input_userName
					+ File.separator + outputSource + File.separator + mailBox() + File.separator
					+ FileNamingUtils.validFileNameForWindows(folderName));
			destinationPath.mkdirs();

			MsOfficeBackup msOffice = new MsOfficeBackup(service, ews.getMapKey(), folderName, destinationPath);

		}

		return outputSource;

	}

	public String downloadImapEmails(String folderPath) {

		if (EmailWizardApplication.r_mbox.isSelected()) {
			String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
			File destinationPath = new File(textField_DownloadingPath.getText() + File.separator + input_userName
					+ File.separator + outputSource + File.separator + validateImapFolderName(changeDelimeter));
			destinationPath.mkdirs();

			MboxrdStorageWriter mbox = new MboxrdStorageWriter(
					destinationPath.getAbsolutePath() + File.separator + destinationPath.getName() + ".mbx", false);
			ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
					iconnforimap_input, mbox, destinationPath, folderPath);
			imapEmailBackUp.downloadImapEmails();
			mbox.close();

		} else if (EmailWizardApplication.r_pst.isSelected()) {
			String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
			String labelBackwordSlash = validateImapFolderName(changeDelimeter).replace("/", "\\");
			pstfolderInfo = pst.getRootFolder().addSubFolder(labelBackwordSlash, true);
			ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
					iconnforimap_input, pst, pstfolderInfo, folderPath);
			imapEmailBackUp.downloadImapEmails();

		} else if (r_office.isSelected() || r_aol.isSelected() || r_aws.isSelected() || r_gmail.isSelected()
				|| r_yahoo.isSelected() || r_yandex.isSelected() || r_zoho.isSelected() || r_icloud.isSelected()
				|| r_hotmail.isSelected()) {

			if (clientforimap_input.existFolder(folderPath)) {
				String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
				String[] split = validateImapFolderName(changeDelimeter).split("/");
				String parentPath = imapFolderPath;
				for (String string : split) {
					String folderPathImap = parentPath + "/" + string;

					if (!clientforimap_Output.existFolder(iconnforimap_Output, folderPathImap)) {
						clientforimap_Output.createFolder(iconnforimap_Output, folderPathImap);
					}
					parentPath = parentPath + "/" + string;
					clientforimap_Output.selectFolder(iconnforimap_Output, folderPathImap);
//					if(r_aws.isSelected()){
//						clientforimap_Output.selectFolder(iconnforimap_Output, folderPathImap);
//					}
//					else
//					{
//						clientforimap_Output.selectFolder(iconnforimap_Output, folderPathImap);
//						clientforimap_Output.subscribeFolder(iconnforimap_Output, folderPathImap);
//					}
				}

				ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
						iconnforimap_input, clientforimap_Output, iconnforimap_Output, parentPath, folderPath);

				imapEmailBackUp.downloadImapEmails();
			}

		} else if (r_imap.isSelected() || r_hostgator.isSelected()) {

			if (clientforimap_input.existFolder(folderPath)) {
				String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
				String[] split = validateImapFolderName(changeDelimeter).split("/");
				String parentPath = imapFolderPath;
				String removeDotlabelName = null;
				String removeSalshWithDotlabelName = null;
				for (String string : split) {

					removeDotlabelName = string.replace(".", "-");
					removeSalshWithDotlabelName = removeDotlabelName.replace("/", ".");

					String folderPathImap = parentPath + "." + removeSalshWithDotlabelName;
					if (!clientforimap_Output.existFolder("INBOX." + folderPathImap)) {
						clientforimap_Output.createFolder(iconnforimap_Output, "INBOX." + folderPathImap);
					}
					removeDotlabelName = string.replace(".", "-");
					removeSalshWithDotlabelName = removeDotlabelName.replace("/", ".");

					parentPath = parentPath + "." + removeSalshWithDotlabelName;
					clientforimap_Output.selectFolder(iconnforimap_Output, "INBOX." + folderPathImap);
					clientforimap_Output.subscribeFolder(iconnforimap_Output, "INBOX." + folderPathImap);
				}
				ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
						iconnforimap_input, clientforimap_Output, iconnforimap_Output, parentPath, folderPath);

				imapEmailBackUp.downloadImapEmails();
			}

		} else if (r_gmail_app.isSelected()) {

			if (clientforimap_input.existFolder(folderPath)) {
				String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
				String mainAccountUserName = new String(InputtxtUserName.getText()).trim();
				String[] split = validateImapFolderName(changeDelimeter).split("/");
				String parentPath = imapFolderPath;
				for (String string : split) {

					try {
						String folderPathImap = parentPath + "/" + string;

						Label label = new Label();
						label.setName(folderPathImap);
						label.setLabelListVisibility("labelShow");
						label.setMessageListVisibility("show");
						if (!lableList.contains(label)) {
							parentLabel = outputGmailService.users().labels().create(user, label)
									.execute();
							lableList.add(label);
						}

					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}

					parentPath = parentPath + "/" + string;
				}

				ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
						iconnforimap_input, mainAccountUserName, parentLabel, outputGmailService, parentPath,
						folderPath);

				imapEmailBackUp.downloadImapEmails();
			}

		}

		else if (EmailWizardApplication.r_csv.isSelected()) {
			String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
			File destinationPath = new File(textField_DownloadingPath.getText() + File.separator + input_userName
					+ File.separator + outputSource + File.separator + validateImapFolderName(changeDelimeter));
			destinationPath.mkdirs();

			Workbook workbook = null;
			ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
					iconnforimap_input, destinationPath, folderPath, workbook);

			try {
				imapEmailBackUp.downloadImapEmails();
			} finally {
				workbook = imapEmailBackUp.saveCSV(destinationPath);
				workbook.dispose();
			}

		}

		else {

			String changeDelimeter = FileNamingUtils.changeImapFolderNameDelimeter(clientforimap_input, folderPath);
			File destinationPath = new File(textField_DownloadingPath.getText() + File.separator + input_userName
					+ File.separator + outputSource + File.separator + validateImapFolderName(changeDelimeter));
			destinationPath.mkdirs();
			ImapEmailBackUp imapEmailBackUp = new ImapEmailBackUp(selectedInput, clientforimap_input,
					iconnforimap_input, destinationPath, folderPath);
			imapEmailBackUp.downloadImapEmails();

		}

		return outputSource;

	}

	public String validateImapFolderName(String folderPath) {
		if (selectedInput.equals(InputSource.IMAP.getValue())
				|| selectedInput.equals(InputSource.HOSTGATOR.getValue())) {
			folderPath = folderPath.replace("INBOX.", "").replace(".", "/");

		}

		return FileNamingUtils.validFileNameForWindows(folderPath);
	}

	public String mailBox() {
		String mailBox = null;
		if (r_archive.isSelected()) {
			mailBox = "Archive Folder";

		} else if (r_public.isSelected()) {
			mailBox = "Public Folder";

		} else if (r_mailbox.isSelected()) {
			mailBox = "Maibox Folder";

		}
		return mailBox;
	}

	public void buttonDisables() {
		count = 0;
		rownCount = 0;
		dataName.setText("");
		EmailWizardApplication.downloadingFileName.setText("");
		progressBar_Downloading.setVisible(true);
		InputtxtUserName.setEnabled(false);
		txtServiceAccountIDorImapHostName.setEnabled(false);
		textField_p12FileAndPortNo.setEnabled(false);
		textField_DownloadingPath.setEnabled(false);
		btnBrowseP2File.setEnabled(false);
		btn_Login.setEnabled(false);
		btnBack_p2.setEnabled(false);
		btnNext_p2.setEnabled(false);
		btnBack_p3.setEnabled(false);
		btnNew_p3.setEnabled(false);
		chckbx_Proxy.setEnabled(false);
		btnDownloadingPath.setEnabled(false);
		textField_DownloadingPath.setEnabled(false);
		btnDownloading.setEnabled(false);
		btnBack_p4.setEnabled(false);
		label_LoginGif.setVisible(true);
		lblDownloading.setVisible(true);
		btnStop.setVisible(true);
		progressBar_Downloading.setVisible(true);
		ProgressBarPanel.setVisible(true);
		stop = false;

	}

	void openBrowser(String url) {
		if (Desktop.isDesktopSupported()) {
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.browse(new URI(url));
			} catch (IOException | URISyntaxException e) {
				logger.warn("Warning : " + e.getMessage());
			}
		} else {
			Runtime runtime = Runtime.getRuntime();
			try {
				runtime.exec("xdg-open " + url);
			} catch (IOException e) {
				logger.warn("Warning : " + e.getMessage());
			}
		}
	}

	public void getImapFolders(ImapFolderInfoCollection imapFolderInfoCollection) {

		for (ImapFolderInfo imapFolderInfo : imapFolderInfoCollection) {

			if (imapFolderInfo.hasChildren()) {
				try {
					ImapFolderInfoCollection folderInfoColl = clientforimap_input.listFolders(iconnforimap_input,imapFolderInfo.getName());
					int totalMessages=returnTotalMessageCount(imapFolderInfo.getName());
					model.addRow(new Object[] { count, imapFolderInfo.getName(), totalMessages, true });
					count++;
					getImapFolders(folderInfoColl);
				} catch (Exception exception) {
					exception.printStackTrace();
					logger.error("Error : " + exception.getMessage());
				}

			} else {
				try {
						
					int totalMessages = returnTotalMessageCount(imapFolderInfo.getName());	
					model.addRow(new Object[] { count, imapFolderInfo.getName(), totalMessages, true });
					count++;
				} catch (Exception exception) {
					exception.printStackTrace();
					logger.error("Error : " + exception.getMessage());
				}

			}

		}
	}
	
	public int returnTotalMessageCount(String folderName)
	{
		ImapFolderInfo folderinfo = clientforimap_input.getFolderInfo(iconnforimap_input, folderName);
		if(Objects.nonNull(folderinfo))
		{
			return folderinfo.getTotalMessageCount();
		}
		return 0;
	}

	public void isHostAndPortNoSelected() {
		if (chckbx_Proxy.isSelected()) {
			lblServiceaccountIdAndHostName.setText("Host Name");
			lblServiceaccountIdAndHostName.setVisible(true);
			txtServiceAccountIDorImapHostName.setVisible(true);
			lblPFileAndPortNumber.setText("Port No.");
			textField_p12FileAndPortNo.setVisible(true);
			lblPFileAndPortNumber.setVisible(true);
		}

	}

	public void selectInputSource(String selectedInput) {

		LogUtils.setTextToLogScreen(textPane_log, logger, "Selected Input Source  : " + selectedInput);

		CardLayout card = (CardLayout) LoginPanel_1.getLayout();
		card.show(LoginPanel_1, "panel_login");

		if (InputSource.GMAIL.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.gmail.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		}
		if (InputSource.AWS.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.mail.us-east-1.awsapps.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		}

		if (InputSource.GMAIL_APP.getValue().equals(selectedInput)) {
			inputIMAPHostName = "imap.gmail.com";
			inputIMAPPortNo = 993;
			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

			chckbx_Proxy.setVisible(false);
			lblServiceaccountIdAndHostName.setVisible(false);
			txtServiceAccountIDorImapHostName.setVisible(false);
			txtServiceAccountIDorImapHostName.setText("");
			textField_p12FileAndPortNo.setText("");

			textField_p12FileAndPortNo.setVisible(false);
			lblPFileAndPortNumber.setVisible(false);
			passwordField_1.setVisible(false);
			lblPassworduser.setVisible(false);
			l_contact.setVisible(true);
			l_calendar.setVisible(true);
			l_drive.setVisible(true);
			l_photos.setVisible(true);
			c_calendar.setVisible(true);
			c_contact.setVisible(true);
			c_drive.setVisible(true);
			c_photos.setVisible(true);

		} else if (InputSource.IMAP.getValue().equals(selectedInput)) {

			ButtonActionImap("", 993);
			textField_p12FileAndPortNo.setText(String.valueOf(""));
			chckbx_Proxy.setVisible(false);
			lblServiceaccountIdAndHostName.setText("Host Name");
			lblServiceaccountIdAndHostName.setVisible(true);
			txtServiceAccountIDorImapHostName.setVisible(true);

			lblPFileAndPortNumber.setText("Port No.");
			textField_p12FileAndPortNo.setVisible(true);
			lblPFileAndPortNumber.setVisible(true);

		} else if (InputSource.HOSTGATOR.getValue().equals(selectedInput)) {

			ButtonActionImap("", 993);
			textField_p12FileAndPortNo.setText(String.valueOf(""));
			chckbx_Proxy.setVisible(false);
			lblServiceaccountIdAndHostName.setText("Host Name");
			lblServiceaccountIdAndHostName.setVisible(true);
			txtServiceAccountIDorImapHostName.setVisible(true);

			lblPFileAndPortNumber.setText("Port No.");
			textField_p12FileAndPortNo.setVisible(true);
			lblPFileAndPortNumber.setVisible(true);

		}

		else if (InputSource.YAHOO.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.mail.yahoo.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);
		} else if (InputSource.Office365.getValue().equals(selectedInput)) {

			inputIMAPHostName = "outlook.office365.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.AOL.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.aol.com";
			inputIMAPPortNo = 993;
			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);
		} else if (InputSource.HOTMAIL.getValue().equals(selectedInput)) {
			inputIMAPHostName = "outlook.office365.com";
			inputIMAPPortNo = 993;
			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.YANDEX.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.yandex.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.ZOHO_EMAIL.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.zoho.com";
			inputIMAPPortNo = 993;
			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.GODADDY.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.secureserver.net";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.ICLOUD.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.mail.me.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.GSUITE.getValue().equals(selectedInput)) {
			chckbx_Proxy.setVisible(false);
			textField_p12FileAndPortNo.setVisible(true);
			lblPFileAndPortNumber.setVisible(true);
			lblServiceaccountIdAndHostName.setVisible(true);
			txtServiceAccountIDorImapHostName.setVisible(true);
			btnBrowseP2File.setVisible(true);
			passwordField_1.setVisible(false);
			lblPassworduser.setVisible(false);

			lblServiceaccountIdAndHostName.setText("Service Account ID");
			lblPFileAndPortNumber.setText("p12 File");

			l_contact.setVisible(true);
			l_calendar.setVisible(true);
			l_drive.setVisible(true);
			l_photos.setVisible(true);
			c_calendar.setVisible(true);
			c_contact.setVisible(true);
			c_drive.setVisible(true);
			c_photos.setVisible(false);

		} else if (InputSource.ONE_AND_ONE_MAIL.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.ionos.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		}

		else if (InputSource.ONE_TWO_SIX.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.126.com";
			inputIMAPPortNo = 993;
			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.ONE_SIX_THREE.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.163.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.AIM.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.aol.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.ARCOR.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.arcor.de";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.ARUBA.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imaps.pec.aruba.it";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		}

		else if (InputSource.ASIA_COM.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.mail.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		}

		else if (InputSource.AT_AND_T.getValue().equals(selectedInput)) {

			inputIMAPHostName = "imap.mail.yahoo.com";
			inputIMAPPortNo = 993;

			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.AXIGEN.getValue().equals(selectedInput)) {

			inputIMAPHostName = "mail.example.com";
			inputIMAPPortNo = 993;
			ButtonActionImap(inputIMAPHostName, inputIMAPPortNo);

		} else if (InputSource.MS_Office_365.getValue().equals(selectedInput)) {
			chckbx_Proxy.setVisible(false);
			passwordField_1.setVisible(false);
			lblPassworduser.setVisible(false);

			lblServiceaccountIdAndHostName.setVisible(false);
			txtServiceAccountIDorImapHostName.setVisible(false);
			textField_p12FileAndPortNo.setVisible(false);
			lblPFileAndPortNumber.setVisible(false);

		} else if (InputSource.Bulk.getValue().equals(selectedInput)) {

			CardLayout loginCardLayout = (CardLayout) LoginPanel_1.getLayout();
			loginCardLayout.show(LoginPanel_1, "panel_TableLogin");

			chckbx_Proxy.setVisible(false);
			passwordField_1.setVisible(false);
			lblPassworduser.setVisible(false);

			lblServiceaccountIdAndHostName.setVisible(false);
			txtServiceAccountIDorImapHostName.setVisible(false);
			textField_p12FileAndPortNo.setVisible(false);
			lblPFileAndPortNumber.setVisible(false);
			l_contact.setVisible(false);
			l_calendar.setVisible(false);
			l_drive.setVisible(false);
			l_photos.setVisible(false);
			c_calendar.setVisible(false);
			c_contact.setVisible(false);
			c_drive.setVisible(false);
			c_photos.setVisible(false);

		}

		else {

			l_contact.setVisible(false);
			l_calendar.setVisible(false);
			l_drive.setVisible(false);
			l_photos.setVisible(false);
			c_calendar.setVisible(false);
			c_contact.setVisible(false);
			c_drive.setVisible(false);
			c_photos.setVisible(false);
		}
	}

	public void ButtonActionImap(String inputIMAPHostName, int inputIMAPPortNo) {

		// txtServiceAccountIDorImapHostName.setText(inputIMAPHostName);
		// textField_p12FileAndPortNo.setText(String.valueOf(inputIMAPPortNo));
		chckbx_Proxy.setVisible(true);
		textField_p12FileAndPortNo.setVisible(false);
		lblPFileAndPortNumber.setVisible(false);
		passwordField_1.setVisible(true);
		btnBrowseP2File.setVisible(false);
		lblPassworduser.setVisible(true);
		lblServiceaccountIdAndHostName.setVisible(false);
		txtServiceAccountIDorImapHostName.setVisible(false);
		c_email.setSelected(true);
		isHostAndPortNoSelected();
	}

	public static ImapClient selectOutputSource() {

		if (r_yahoo.isSelected()) {
			clientforimap_Output = new ImapClient("imap.mail.yahoo.com", 993, output_userName, output_password);
		} else if (r_gmail.isSelected()) {
			clientforimap_Output = new ImapClient("imap.gmail.com", 993, output_userName, output_password);
		} else if (r_icloud.isSelected()) {
			clientforimap_Output = new ImapClient("imap.mail.me.com", 993, output_userName, output_password);
		}

		else if (r_aws.isSelected()) {
			clientforimap_Output = new ImapClient("imap.mail.us-east-1.awsapps.com", 993, output_userName,
					output_password);

		} else if (r_yandex.isSelected()) {
			clientforimap_Output = new ImapClient("imap.yandex.com", 993, output_userName, output_password);
		} else if (r_zoho.isSelected()) {
			clientforimap_Output = new ImapClient("imap.zoho.com", 993, output_userName, output_password);
		} else if (r_hotmail.isSelected()) {
			clientforimap_Output = new ImapClient("outlook.office365.com", 993, output_userName, output_password);
		}

		else if (r_office.isSelected()) {
			clientforimap_Output = new ImapClient("outlook.office365.com", 993, output_userName, output_password);
		} else if (r_aol.isSelected()) {

			clientforimap_Output = new ImapClient("imap.aol.com", 993, output_userName, output_password);
		} else if (r_imap.isSelected() || r_hostgator.isSelected()) {

			clientforimap_Output = new ImapClient(output_imapHost, Integer.valueOf(output_portNo), output_userName,
					output_password);

		}
		return clientforimap_Output;

	}

	public static ImapClient connectionWithInputIMAP() throws Exception {

		clientforimap_input = new ImapClient(inputIMAPHostName, inputIMAPPortNo, input_userName, input_password);
		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);
		EmailClient.setSocketsLayerVersion2(true);

		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_input.setTimeout(60000);
		clientforimap_input.setGreetingTimeout(4000);
		clientforimap_input.setUidPlusSupported(true);

		iconnforimap_input = clientforimap_input.createConnection(true);
		checkImapConnectionTime = System.currentTimeMillis() + IMAP_RERESH_TIMEOUT;
		return clientforimap_input;
	}

	public static ImapClient outputImapConnection() throws Exception {

		clientforimap_Output = selectOutputSource();
		clientforimap_Output.setSecurityOptions(SecurityOptions.Auto);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_Output.setTimeout(100000);
		clientforimap_Output.setGreetingTimeout(6000);
		clientforimap_Output.setUidPlusSupported(true);
		iconnforimap_Output = clientforimap_Output.createConnection();
		checkImapConnectionTime = System.currentTimeMillis() + IMAP_RERESH_TIMEOUT;
		return clientforimap_Output;
	}

	public void outputImapInitialfolderCreation(ImapClient clientforimap_Output) {

		String mainAccountUserName = new String(InputtxtUserName.getText()).trim();
		Calendar calendar = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("dd-MM-yy HH-mm-ss");

		if (r_imap.isSelected() || r_hostgator.isSelected()) {
			imapFolderPath = mainAccountUserName.replace(".", "-") + "-" + dateFormat.format(calendar.getTime());
			clientforimap_Output.createFolder(iconnforimap_Output, "INBOX." + imapFolderPath);
			clientforimap_Output.selectFolder(iconnforimap_Output, "INBOX." + imapFolderPath);
			clientforimap_Output.subscribeFolder(iconnforimap_Output, "INBOX." + imapFolderPath);
		} else {

			imapFolderPath = mainAccountUserName + "-" + dateFormat.format(calendar.getTime());
			clientforimap_Output.createFolder(iconnforimap_Output, imapFolderPath);
			clientforimap_Output.selectFolder(iconnforimap_Output, imapFolderPath);

			clientforimap_Output.subscribeFolder(iconnforimap_Output, imapFolderPath);

		}

	}

	public void outputGmailAppInitialfolderCreation() throws IOException {

		String mainAccountUserName = new String(InputtxtUserName.getText()).trim();
		Calendar calendar = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("dd-MM-yy HH-mm-ss");
		imapFolderPath = mainAccountUserName + "-" + dateFormat.format(calendar.getTime());

		Label label = new Label();
		label.setName(imapFolderPath);
		label.setLabelListVisibility("labelShow");
		label.setMessageListVisibility("show");
		parentLabel = outputGmailService.users().labels().create(user, label).execute();

	}

	public static ExchangeService loginWithRefreshTokenEWS(String username) {

		try {
			String refreshToken = ews.getRefreshToken();
			service = ews.loginRefreshTokenEWS(username, refreshToken);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return service;
	}

	public void setFoldeList(List<String> folderIdlist) {

		this.folderIdlist = folderIdlist;
	}

	public List<String> MsOfficeMigrationTree(List<String> folderIdlist) {

		TreePath[] treePath = folderTree.getCheckingPaths();
		int i = 0;
		for (TreePath tp : treePath) {

			for (Object pathPart : tp.getPath()) {

				if (ews.getMapKeyTree().containsKey(pathPart.toString())) {
					try {
						Folder folder = ews.getMapKeyTree().get(pathPart.toString());
						if (!folderIdlist.contains(folder.getId().toString())) {
							System.out.println(folder.getDisplayName());
							System.out.println(folder.getId().getChangeKey());
							folderIdlist.add(folder.getId().toString());
						}
					} catch (Exception ex) {
						// TODO: handle exception
					}

				}

			}

		}

		return folderIdlist;

	}

	public void MsOfficeMigration() {

		buttonDisables();
		long startTime = new Date().getTime();
		DefaultTableModel dm = (DefaultTableModel) table_Downloading.getModel();
		while (dm.getRowCount() > 0) {
			dm.removeRow(0);
		}
		int rows = table_UserDetails.getRowCount();
		if (r_pst.isSelected()) {
			LogUtils.setTextToLogScreen(textPane_log, logger,
					"Inside Emails Class.." + OutputSource.PST.name() + " Format Selected");

			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
			LocalDateTime now = LocalDateTime.now();
			File destinationPathOfPST = new File(textField_DownloadingPath.getText().trim() + File.separator
					+ input_userName + File.separator + outputSource);
			destinationPathOfPST.mkdirs();
			pst = PersonalStorage.create(destinationPathOfPST.getAbsolutePath() + File.separator + "(" + splitCount
					+ ") " + dtf.format(now) + "-" + input_userName + ".pst", FileFormatVersion.Unicode);
			pstSplitFile = new File(destinationPathOfPST.getAbsolutePath() + File.separator + "(" + splitCount + ") "
					+ dtf.format(now) + "-" + input_userName + ".pst");
			pst.getStore().changeDisplayName(input_userName);
			pstfolderInfo = new FolderInfo();
		}

		for (int i = 0; i < rows; i++) {

			try {
				Object checked = table_UserDetails.getValueAt(i, 3);
				if (!(boolean) checked) {
					continue;
				}
				if (stop) {
					break;
				}
				String folderPath = table_UserDetails.getModel().getValueAt(i, 1).toString();
				LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading Folder : " + folderPath);

				modelDownloading = (DefaultTableModel) EmailWizardApplication.table_Downloading.getModel();
				modelDownloading.addRow(new Object[] { folderPath, 0, 0, 0, 0 });

				if (c_email.isSelected()) {

					dataName.setText("Emails");
					downloadMSOfficeEmails(folderPath);
				}

				rownCount++;
			} catch (ImapException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();

				if (e.getMessage().contains("Unknown Mailbox:")) {
					rownCount++;
					System.out.println(e.getMessage());
				}

			}

			catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				rownCount++;

			}
		}
		long endTime = new Date().getTime() - startTime;
		DateFormat sdf = new SimpleDateFormat("HH-mm-ss");
		sdf.setTimeZone(TimeZone.getTimeZone("UTC"));
		LogUtils.setTextToLogScreen(textPane_log, logger,
				"!!!!!! Backup Complete in (hrs:min:sec) : " + sdf.format(new Date(endTime)) + " !!!!!!!");
		buttonEnables();
	}

	public void GmailAppMigration() {
		buttonDisables();
		lableList = new ArrayList<>();
		long startTime = new Date().getTime();
		modelDownloading = (DefaultTableModel) EmailWizardApplication.table_Downloading.getModel();
		modelDownloading.addRow(new Object[] { 0, 0, 0, 0, 0 });
		File destinationPath = null;

		try {
			if (c_contact.isSelected()) {
				LogUtils.setTextToLogScreen(textPane_log, logger,
						"Downloading Contacts of account : " + input_userName);
				dataName.setText("Contact");
				destinationPath = new File(textField_DownloadingPath.getText() + File.separator + InputtxtUserName.getText()
						+ File.separator + input_userName + File.separator + "Contact");
				destinationPath.mkdirs();
				detinationPath = destinationPath.getAbsolutePath();
				ContactBackup contactBackup = new ContactBackup();
				contactBackup.downloadGMailContact(googleLogin.getoathGoogleCredential());
			}
			if (c_calendar.isSelected()) {

				LogUtils.setTextToLogScreen(textPane_log, logger,
						"Downloading Calendar of account : " + input_userName);
				dataName.setText("Calender");
				destinationPath = new File(textField_DownloadingPath.getText() + File.separator + InputtxtUserName.getText()
						+ File.separator + input_userName + File.separator + "Calender");
				destinationPath.mkdirs();
				detinationPath = destinationPath.getAbsolutePath();
				CalenderBackup calenderBackup = new CalenderBackup();
				calenderBackup.downloadGmailAPPCalendar(googleLogin.getoathGoogleCredential());

			}
			if (c_drive.isSelected()) {

				LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading Drive data of :" + input_userName);
				dataName.setText("Drive");
				destinationPath = new File(textField_DownloadingPath.getText() + File.separator + InputtxtUserName.getText()
						+ File.separator + input_userName + File.separator + "Drive");
				destinationPath.mkdirs();
				detinationPath = destinationPath.getAbsolutePath();
				DriveBackup driveBackup = new DriveBackup();
				driveBackup.googleCredentials(googleLogin.getoathGoogleCredential());

			}
			if (c_photos.isSelected()) {

				LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading photos of account : " + input_userName);
				dataName.setText("Photo");
				destinationPath = new File(textField_DownloadingPath.getText() + File.separator + InputtxtUserName.getText()
						+ File.separator + input_userName + File.separator + "Photos");
				destinationPath.mkdirs();
				detinationPath = destinationPath.getAbsolutePath();
				PhotosBackup photoBackup = new PhotosBackup();
				photoBackup.googleCredentials(googleLogin.getoathGoogleCredentials());

			}
			if (c_email.isSelected()) {

				dataName.setText("Emails");
				GmailBackup gmailBackup = null;

				if (r_office.isSelected() || r_aol.isSelected() || r_aws.isSelected() || r_gmail.isSelected()
						|| r_yahoo.isSelected() || r_yandex.isSelected() || r_zoho.isSelected() || r_icloud.isSelected()
						|| r_hotmail.isSelected() || r_imap.isSelected() || r_hostgator.isSelected()) {

					outputSource = emailClientSelectedFormatAtOutput();
					LogUtils.setTextToLogScreen(textPane_log, logger,"Downloading account : " + input_userName + " in " + outputSource + " format");
						
					String newimapFolderPath = null;
					if (r_imap.isSelected() || r_hostgator.isSelected()) {
						String tempImapPath = imapFolderPath;
						String userName = input_userName.replace(".", "-");
						newimapFolderPath = "INBOX." + tempImapPath + "." + userName;

					} else {
						newimapFolderPath = imapFolderPath + "/" + input_userName;
					}

					clientforimap_Output.createFolder(iconnforimap_Output, newimapFolderPath);

					if (r_aws.isSelected()) {
						clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
					} else {
						clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
						clientforimap_Output.subscribeFolder(iconnforimap_Output, newimapFolderPath);
					}

					gmailBackup = new GmailBackup(googleLogin.getoathGoogleCredential(), InputtxtUserName.getText().trim(),
							null, null, clientforimap_Output, iconnforimap_Output, newimapFolderPath);
					gmailBackup.download();

				}
				 else if (r_gmail_app.isSelected()) {
				outputSource = emailClientSelectedFormatAtOutput();
				LogUtils.setTextToLogScreen(textPane_log, logger,"Downloading account : " + input_userName + " in " + outputSource + " format");
				String mainAccountUserName = new String(InputtxtUserName.getText()).trim();
				String newimapFolderPath = imapFolderPath + "/" + input_userName;
				
				Label label = new Label();
				label.setName(newimapFolderPath);
				label.setLabelListVisibility("labelShow");
				label.setMessageListVisibility("show");
				if (!lableList.contains(label)) {
					parentLabel = outputGmailService.users().labels().create(user, label).execute();							
					lableList.add(label);
				}
				
				gmailBackup = new GmailBackup(inputGmailCredential, InputtxtUserName.getText().trim(),
						null, null, mainAccountUserName, parentLabel, outputGmailService, newimapFolderPath);
				   gmailBackup.download();
				
				
			 }
				else if (r_csv.isSelected()) {
					LogUtils.setTextToLogScreen(textPane_log, logger,
							"Downloading Gmail of account : " + input_userName + " in " + outputSource + " format");
					destinationPath = new File(textField_DownloadingPath.getText() + File.separator
							+ InputtxtUserName.getText() + File.separator + input_userName + File.separator + "Emails"
							+ File.separator + outputSource);
					destinationPath.mkdirs();
					detinationPath = destinationPath.getAbsolutePath();
					Workbook workbook = null;

					gmailBackup = new GmailBackup(googleLogin.getoathGoogleCredential(), null, null, null,
							detinationPath, workbook);
					gmailBackup.download();

				} else {
					LogUtils.setTextToLogScreen(textPane_log, logger,
							"Downloading Gmail of account : " + input_userName + " in " + " format");

					destinationPath = new File(textField_DownloadingPath.getText() + File.separator
							+ InputtxtUserName.getText() + File.separator + input_userName + File.separator + "Emails"
							+ File.separator + outputSource);
					destinationPath.mkdirs();
					detinationPath = destinationPath.getAbsolutePath();
					gmailBackup = new GmailBackup(googleLogin.getoathGoogleCredential(), InputtxtUserName.getText().trim(),
							null, null, detinationPath);
					gmailBackup.download();
				}

			}
		} catch (GeneralSecurityException | IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}

		long endTime = new Date().getTime() - startTime;
		DateFormat sdf = new SimpleDateFormat("HH-mm-ss");
		sdf.setTimeZone(TimeZone.getTimeZone("UTC"));
		LogUtils.setTextToLogScreen(textPane_log, logger,
				"!!!!!! Backup Complete in (hrs:min:sec) : " + sdf.format(new Date(endTime)) + " !!!!!!!");
		buttonEnables();

	}

	public void GSuiteMigration() {
		buttonDisables();
		lableList = new ArrayList<>();
		long startTime = new Date().getTime();
		rownCount = 0;
		DefaultTableModel dm = (DefaultTableModel) table_Downloading.getModel();
		while (dm.getRowCount() > 0) {
			dm.removeRow(0);
		}
		int rows = table_UserDetails.getRowCount();
		File destinationPath = null;
		for (int i = 0; i < rows; i++) {

			try {

				Object checked = table_UserDetails.getValueAt(i, 3);
				if (!(boolean) checked) {
					continue;
				}
				if (stop) {
					break;
				}

				String value = table_UserDetails.getModel().getValueAt(i, 2).toString();
				modelDownloading = (DefaultTableModel) EmailWizardApplication.table_Downloading.getModel();
				modelDownloading.addRow(new Object[] { value, 0, 0, 0, 0 });

				if (c_contact.isSelected()) {
					LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading Contacts of account : " + value);
					dataName.setText("Contact");

					destinationPath = new File(textField_DownloadingPath.getText() + File.separator
							+ InputtxtUserName.getText() + File.separator + value + File.separator + "Contact");
					destinationPath.mkdirs();
					detinationPath = destinationPath.getAbsolutePath();
					ContactBackup contactBackup = new ContactBackup();
					contactBackup.downloadGsuiteContact(txtServiceAccountIDorImapHostName.getText(), value,
							textField_p12FileAndPortNo.getText());

				}

				if (c_drive.isSelected()) {

					LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading Drive data of :" + value);
					dataName.setText("Drive");
					destinationPath = new File(textField_DownloadingPath.getText() + File.separator
							+ InputtxtUserName.getText() + File.separator + value + File.separator + "Drive");
					destinationPath.mkdirs();
					detinationPath = destinationPath.getAbsolutePath();
					DriveBackup driveBackup = new DriveBackup();

					driveBackup.googleCredentials(txtServiceAccountIDorImapHostName.getText(), value,
							textField_p12FileAndPortNo.getText());

				}
				if (c_photos.isSelected()) {

					LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading photos of account : " + value);
					dataName.setText("Photo");

					destinationPath = new File(textField_DownloadingPath.getText() + File.separator
							+ InputtxtUserName.getText() + File.separator + value + File.separator + "Photos");
					destinationPath.mkdirs();

					detinationPath = destinationPath.getAbsolutePath();
					PhotosBackup photoBackup = new PhotosBackup();

					photoBackup.googleCredentials(txtServiceAccountIDorImapHostName.getText(), value,
							textField_p12FileAndPortNo.getText());

				}
				if (c_calendar.isSelected()) {

					LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading Calendar of account : " + value);
					dataName.setText("Calender");
					destinationPath = new File(textField_DownloadingPath.getText() + File.separator
							+ InputtxtUserName.getText() + File.separator + value + File.separator + "Calender");
					destinationPath.mkdirs();
					detinationPath = destinationPath.getAbsolutePath();
					CalenderBackup calenderBackup = new CalenderBackup();

					calenderBackup.downloadGsuiteCalendar(txtServiceAccountIDorImapHostName.getText(), value,
							textField_p12FileAndPortNo.getText());

				}

				if (c_email.isSelected()) {

					String format = outputSelectedFileFormat();
					dataName.setText("Emails");
					GmailBackup gmailBackup = null;

					if (r_office.isSelected() || r_aol.isSelected() || r_aws.isSelected() || r_gmail.isSelected()
							|| r_yahoo.isSelected() || r_yandex.isSelected() || r_zoho.isSelected()
							|| r_icloud.isSelected() || r_hotmail.isSelected() || r_imap.isSelected()
							|| r_hostgator.isSelected()) {

						format = emailClientSelectedFormatAtOutput();
						LogUtils.setTextToLogScreen(textPane_log, logger,
								"Downloading  account : " + value + " in " + format + " format");

						String newimapFolderPath = null;
						if (r_imap.isSelected() || r_hostgator.isSelected()) {
							String tempImapPath = imapFolderPath;
							String userName = value.replace(".", "-");
							newimapFolderPath = "INBOX." + tempImapPath + "." + userName;
						} else {
							newimapFolderPath = imapFolderPath + "/" + value;
						}

						clientforimap_Output.createFolder(iconnforimap_Output, newimapFolderPath);

						if (r_aws.isSelected()) {
							clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
						} else {
							clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
							clientforimap_Output.subscribeFolder(iconnforimap_Output, newimapFolderPath);
						}

						gmailBackup = new GmailBackup(null, value, txtServiceAccountIDorImapHostName.getText(),
								textField_p12FileAndPortNo.getText(), clientforimap_Output, iconnforimap_Output,
								newimapFolderPath);
						gmailBackup.download();

					} 
					 else if (r_gmail_app.isSelected()) {

						    format = emailClientSelectedFormatAtOutput();
							LogUtils.setTextToLogScreen(textPane_log, logger,
									"Downloading  account : " + value + " in " + format + " format");
							String mainAccountUserName = new String(InputtxtUserName.getText()).trim();
							String newimapFolderPath =  imapFolderPath + "/" + value;
							
							Label label = new Label();
							label.setName(newimapFolderPath);
							label.setLabelListVisibility("labelShow");
							label.setMessageListVisibility("show");
							if (!lableList.contains(label)) {
								parentLabel = outputGmailService.users().labels().create(user, label).execute();							
								lableList.add(label);
							}
							   gmailBackup = new GmailBackup(null, value, txtServiceAccountIDorImapHostName.getText(),
										textField_p12FileAndPortNo.getText(), mainAccountUserName, parentLabel, outputGmailService,
										newimapFolderPath);
								gmailBackup.download();							
							
						 }
					
									
					else if (r_csv.isSelected()) {
						LogUtils.setTextToLogScreen(textPane_log, logger,
								"Downloading Gmail of account : " + value + " in " + format + " format");
						destinationPath = new File(
								textField_DownloadingPath.getText() + File.separator + InputtxtUserName.getText()
										+ File.separator + value + File.separator + "Emails" + File.separator + format);
						destinationPath.mkdirs();
						detinationPath = destinationPath.getAbsolutePath();
						Workbook workbook = null;

						gmailBackup = new GmailBackup(null, value, txtServiceAccountIDorImapHostName.getText(),
								textField_p12FileAndPortNo.getText(), detinationPath, workbook);
						gmailBackup.download();

					} else {
						LogUtils.setTextToLogScreen(textPane_log, logger,
								"Downloading Gmail of account : " + value + " in " + format + " format");
						destinationPath = new File(
								textField_DownloadingPath.getText() + File.separator + InputtxtUserName.getText()
										+ File.separator + value + File.separator + "Emails" + File.separator + format);
						destinationPath.mkdirs();
						detinationPath = destinationPath.getAbsolutePath();
						gmailBackup = new GmailBackup(null, value, txtServiceAccountIDorImapHostName.getText(),
								textField_p12FileAndPortNo.getText(), detinationPath);
						gmailBackup.download();
					}

				}

				rownCount++;
			} catch (GeneralSecurityException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
			}

		}
		long endTime = new Date().getTime() - startTime;
		DateFormat sdf = new SimpleDateFormat("HH-mm-ss");
		sdf.setTimeZone(TimeZone.getTimeZone("UTC"));
		LogUtils.setTextToLogScreen(textPane_log, logger,
				"!!!!!! Backup Complete in (hrs:min:sec) : " + sdf.format(new Date(endTime)) + " !!!!!!!");
		buttonEnables();
	}

	public void downloadImapFoldersForBulk(ImapFolderInfoCollection imapFolderInfoCollection) {

		for (ImapFolderInfo imapFolderInfo : imapFolderInfoCollection) {

			if (imapFolderInfo.hasChildren()) {

				ImapFolderInfoCollection folderInfoColl = clientforimap_input.listFolders(imapFolderInfo.getName());
				int totalMessages = imapFolderInfo.getTotalMessageCount();
				if (c_email.isSelected()) {

					LogUtils.setTextToLogScreen(textPane_log, logger,
							"Downloading Folder : " + imapFolderInfo.getName() + " Total Message :" + totalMessages);
					dataName.setText("Emails");
					try {
						downloadImapEmails(imapFolderInfo.getName());
					} catch (ImapException e) {
						e.printStackTrace();

						if (e.getMessage().contains("Unknown Mailbox:")) {
							logger.warn("Unknown Mailbox : " + e.getMessage());
						}

					}

				}
				downloadImapFoldersForBulk(folderInfoColl);

			} else {
				ImapFolderInfo folderinfo = clientforimap_input.getFolderInfo(iconnforimap_input,
						imapFolderInfo.getName());
				int totalMessages = folderinfo.getTotalMessageCount();

				if (c_email.isSelected()) {
					LogUtils.setTextToLogScreen(textPane_log, logger,
							"Downloading Folder : " + folderinfo + " Total Message :" + totalMessages);
					dataName.setText("Emails");
					try {
						downloadImapEmails(imapFolderInfo.getName());
					} catch (ImapException e) {
						e.printStackTrace();

						if (e.getMessage().contains("Unknown Mailbox:")) {

							logger.warn("Unknown Mailbox : " + e.getMessage());
						}

					}
				}

			}

		}
	}

	public void ImapBulkMigration() {
		DefaultTableModel dm = (DefaultTableModel) table_Login.getModel();
		modelDownloading = (DefaultTableModel) EmailWizardApplication.table_Downloading.getModel();
		buttonDisables();
		String downloadingPath = textField_DownloadingPath.getText();
		textField_DownloadingPath.setText(downloadingPath + File.separator + ToolDetails.messageboxtitle);
		long startTime = new Date().getTime();
		for (int i = 0; i < dm.getRowCount(); i++) {
			try {

				if (stop) {
					break;
				}
				count = 0;
				input_userName = table_Login.getModel().getValueAt(i, 1).toString().trim();
				input_password = table_Login.getModel().getValueAt(i, 2).toString().trim();
				inputIMAPHostName = table_Login.getModel().getValueAt(i, 3).toString().trim();
				inputIMAPPortNo = Integer.parseInt(table_Login.getModel().getValueAt(i, 4).toString().trim());
				modelDownloading.addRow(new Object[] { input_userName, 0, 0, 0, 0 });
				LogUtils.setTextToLogScreen(textPane_log, logger, "Connecting with : " + i + " " + input_userName);
				loginTableModel.setValueAt("Connecting..", i, 5);
				modelDownloading.setValueAt("Connecting..Please Wait", rownCount, 2);

				clientforimap_input = connectionWithInputIMAP();// Imap connection
				clientforimap_input.setUseMultiConnection(MultiConnectionMode.Enable);

				loginTableModel.setValueAt("Connected", i, 5);
				modelDownloading.setValueAt("Connected", EmailWizardApplication.rownCount, 2);

				LogUtils.setTextToLogScreen(textPane_log, logger, "Connection done with : " + i + " " + input_userName);
				modelDownloading.setValueAt("In Progress", rownCount, 2);// imap bulk download
				if (r_pst.isSelected()) {
					DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
					LocalDateTime now = LocalDateTime.now();
					File destinationPathOfPST = new File(textField_DownloadingPath.getText().trim() + File.separator
							+ input_userName + File.separator + outputSource);
					destinationPathOfPST.mkdirs();
					pst = PersonalStorage.create(destinationPathOfPST.getAbsolutePath() + File.separator + "(" + 0
							+ ") " + dtf.format(now) + "-" + input_userName + ".pst", FileFormatVersion.Unicode);
					pstSplitFile = new File(destinationPathOfPST.getAbsolutePath() + File.separator + "(" + 0 + ") "
							+ dtf.format(now) + "-" + input_userName + ".pst");
					pst.getStore().changeDisplayName(input_userName);
					pstfolderInfo = new FolderInfo();
				}

				ImapFolderInfoCollection folderInfoColl = clientforimap_input.listFolders();
				downloadImapFoldersForBulk(folderInfoColl);

				modelDownloading.setValueAt("Completed", rownCount, 2);
				rownCount++;

			} catch (Exception ex) {
				TableColumn tColumn = table_Downloading.getColumnModel().getColumn(2);
				tColumn.setCellRenderer(new PaintTableCellRenderer(Color.RED));
				logger.error("Error Found ", ex);
				modelDownloading.setValueAt("Connection Not Estalished", rownCount, 2);
				rownCount++;
			} finally {
				if (pst != null) {
					pst.close();
				}
			}

		}
		textField_DownloadingPath.setText(downloadingPath);
		long endTime = new Date().getTime() - startTime;
		DateFormat sdf = new SimpleDateFormat("HH-mm-ss");
		sdf.setTimeZone(TimeZone.getTimeZone("UTC"));
		LogUtils.setTextToLogScreen(textPane_log, logger,
				"!!!!!! Backup Complete in (hrs:min:sec) : " + sdf.format(new Date(endTime)) + " !!!!!!!");
		buttonEnables();

	}

	public void ImapMigration() {
		lableList = new ArrayList<>();
		buttonDisables();
		long startTime = new Date().getTime();
		DefaultTableModel dm = (DefaultTableModel) table_Downloading.getModel();
		while (dm.getRowCount() > 0) {
			dm.removeRow(0);
		}
		int rows = table_UserDetails.getRowCount();
		if (r_pst.isSelected()) {
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
			LocalDateTime now = LocalDateTime.now();
			File destinationPathOfPST = new File(textField_DownloadingPath.getText().trim() + File.separator
					+ input_userName + File.separator + outputSource);
			destinationPathOfPST.mkdirs();
			pst = PersonalStorage.create(destinationPathOfPST.getAbsolutePath() + File.separator + "(" + 0 + ") "
					+ dtf.format(now) + "-" + input_userName + ".pst", FileFormatVersion.Unicode);
			pstSplitFile = new File(destinationPathOfPST.getAbsolutePath() + File.separator + "(" + 0 + ") "
					+ dtf.format(now) + "-" + input_userName + ".pst");
			pst.getStore().changeDisplayName(input_userName);
			pstfolderInfo = new FolderInfo();
		}
		for (int i = 0; i < rows; i++) {

			try {
				Object checked = table_UserDetails.getValueAt(i, 3);
				if (!(boolean) checked) {
					continue;
				}
				if (stop) {
					break;
				}

				String imapfolderPath = table_UserDetails.getModel().getValueAt(i, 1).toString();
				modelDownloading = (DefaultTableModel) EmailWizardApplication.table_Downloading.getModel();
				modelDownloading.addRow(new Object[] { imapfolderPath, 0, 0, 0, 0 });

				if (c_email.isSelected()) {

					LogUtils.setTextToLogScreen(textPane_log, logger, "Downloading Folder : " + imapfolderPath);
					dataName.setText("Emails");

					downloadImapEmails(imapfolderPath);

					TableColumn tColumn = table_Downloading.getColumnModel().getColumn(3);
					tColumn.setCellRenderer(new ColumnColorRenderer(Color.green, Color.BLACK));
				}
				rownCount++;
			} catch (ImapException e) {
				if (e.getMessage().contains("Unknown Mailbox:")
						|| e.getMessage().contains("AE_4_2_0181 NO [CANNOT] Folder name is not allowed.")) {

					logger.warn("Unknown Mailbox : " + e.getMessage());
				} else {
					e.printStackTrace();
				}
				rownCount++;

			}

			catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				rownCount++;
			}
		}

		long endTime = new Date().getTime() - startTime;
		DateFormat sdf = new SimpleDateFormat("HH-mm-ss");
		sdf.setTimeZone(TimeZone.getTimeZone("UTC"));
		LogUtils.setTextToLogScreen(textPane_log, logger,
				"!!!!!! Backup Complete in (hrs:min:sec) : " + sdf.format(new Date(endTime)) + " !!!!!!!");

		buttonEnables();
	}

	public String outputSelectedFileFormat() {
		if (EmailWizardApplication.r_Eml.isSelected()) {
			return OutputSource.EML.name();
		} else if (EmailWizardApplication.r_pdf.isSelected()) {
			return OutputSource.PDF.name();
		} else if (EmailWizardApplication.r_pst.isSelected()) {
			return OutputSource.PST.name();
		} else if (EmailWizardApplication.r_msg.isSelected()) {
			return OutputSource.MSG.name();
		} else if (EmailWizardApplication.r_emlx.isSelected()) {
			return OutputSource.EMLX.name();
		} else if (EmailWizardApplication.r_mbox.isSelected()) {
			return OutputSource.MBOX.name();
		} else if (EmailWizardApplication.r_html.isSelected()) {
			return OutputSource.HTML.name();
		} else if (EmailWizardApplication.r_rtf.isSelected()) {
			return OutputSource.RTF.name();
		} else if (EmailWizardApplication.r_xps.isSelected()) {
			return OutputSource.XPS.name();
		} else if (EmailWizardApplication.r_emf.isSelected()) {
			return OutputSource.EMF.name();
		} else if (EmailWizardApplication.r_docx.isSelected()) {
			return OutputSource.DOCX.name();
		} else if (EmailWizardApplication.r_jpeg.isSelected()) {
			return OutputSource.JPEG.name();
		} else if (EmailWizardApplication.r_docm.isSelected()) {
			return OutputSource.DOCM.name();
		} else if (EmailWizardApplication.r_docm.isSelected()) {
			return OutputSource.DOCM.name();
		} else if (EmailWizardApplication.r_text.isSelected()) {
			return OutputSource.TEXT.name();
		} else if (EmailWizardApplication.r_png.isSelected()) {
			return OutputSource.PNG.name();
		} else if (EmailWizardApplication.r_tiff.isSelected()) {
			return OutputSource.TIFF.name();
		} else if (EmailWizardApplication.r_svg.isSelected()) {
			return OutputSource.SVG.name();
		} else if (EmailWizardApplication.r_epub.isSelected()) {
			return OutputSource.EPUB.name();
		} else if (EmailWizardApplication.r_dotm.isSelected()) {
			return OutputSource.DOTM.name();
		} else if (EmailWizardApplication.r_bmp.isSelected()) {
			return OutputSource.BMP.name();
		} else if (EmailWizardApplication.r_gif.isSelected()) {
			return OutputSource.GIF.name();
		} else if (EmailWizardApplication.r_ott.isSelected()) {
			return OutputSource.OTT.name();
		} else if (EmailWizardApplication.r_wordml.isSelected()) {
			return OutputSource.WORLD_ML.name();
		}

		else if (EmailWizardApplication.r_odt.isSelected()) {
			return OutputSource.ODT.name();
		} else if (r_csv.isSelected()) {
			return OutputSource.CSV.name();
		}

		return "No file format Selected";
	}

	public String emailClientSelectedFormatAtOutput() {
		if (EmailWizardApplication.r_aol.isSelected()) {
			return OutputSource.AOL.name();
		} else if (EmailWizardApplication.r_aws.isSelected()) {
			return OutputSource.AWS.name();
		} else if (EmailWizardApplication.r_icloud.isSelected()) {
			return OutputSource.ICLOUD.name();
		} else if (EmailWizardApplication.r_gmail.isSelected()) {
			return OutputSource.GMAIL.name();
		} else if (EmailWizardApplication.r_office.isSelected()) {
			return OutputSource.Office365.name();
		} else if (EmailWizardApplication.r_yahoo.isSelected()) {
			return OutputSource.YAHOO.name();
		} else if (EmailWizardApplication.r_zoho.isSelected()) {
			return OutputSource.ZOHO_EMAIL.name();
		} else if (r_hotmail.isSelected()) {
			return OutputSource.Hotmail.name();
		} else if (EmailWizardApplication.r_yandex.isSelected()) {
			return OutputSource.YANDEX.name();
		} else if (EmailWizardApplication.r_imap.isSelected()) {
			return OutputSource.IMAP.name();
		} else if (EmailWizardApplication.r_hostgator.isSelected()) {
			return OutputSource.HostGator.name();
		} else if (EmailWizardApplication.r_gmail_app.isSelected()) {			
				
			return OutputSource.GMAIL_APP.name();
		}
		return "No email Client selected ";

	}
	public void changeHeader() {
		JTableHeader th = table_UserDetails.getTableHeader();
		TableColumnModel tcm = th.getColumnModel();

		if (selectedInput.equals(InputSource.GSUITE.getValue())) {
			TableColumn tc = tcm.getColumn(0);
			tc.setHeaderValue("<HTML><B>User No</B></HTML>");
			tc = tcm.getColumn(1);
			tc.setHeaderValue("<HTML><B>Name</B></HTML>");
			tc = tcm.getColumn(2);
			tc.setHeaderValue("<HTML><B>User Email Address</B></HTML>");

		} else {
			TableColumn tc = tcm.getColumn(0);
			tc.setHeaderValue("<HTML><B>Folder(s). No</B></HTML>");
			tc = tcm.getColumn(1);
			tc.setHeaderValue("<HTML><B>Folders Name</B></HTML>");
			tc = tcm.getColumn(2);
			tc.setHeaderValue("<HTML><B>Count</B></HTML>");
		}

		th.repaint();
	}

	public void changeHeaderoutput() {
		JTableHeader th = table_Downloading.getTableHeader();
		TableColumnModel tcm = th.getColumnModel();

		if (selectedInput.equals(InputSource.GSUITE.getValue())) {
			TableColumn tc = tcm.getColumn(0);
			tc.setHeaderValue("<HTML><B>User Name</B></HTML>");
			tc = tcm.getColumn(1);
			tc.setHeaderValue("<HTML><B>Drive</B></HTML>");
			tc = tcm.getColumn(2);
			tc.setHeaderValue("<HTML><B>Contact</B></HTML>");
			tc = tcm.getColumn(3);
			tc.setHeaderValue("<HTML><B>Calendar</B></HTML>");
			tc = tcm.getColumn(4);
			tc.setHeaderValue("<HTML><B>Gmail</B></HTML>");
		} else if (selectedInput.equals(InputSource.GMAIL_APP.getValue())) {
			TableColumn tc = tcm.getColumn(0);
			tc.setHeaderValue("<HTML><B>Photos</B></HTML>");
			tc = tcm.getColumn(1);
			tc.setHeaderValue("<HTML><B>Drive</B></HTML>");
			tc = tcm.getColumn(2);
			tc.setHeaderValue("<HTML><B>Contact</B></HTML>");
			tc = tcm.getColumn(3);
			tc.setHeaderValue("<HTML><B>Calendar</B></HTML>");
			tc = tcm.getColumn(4);
			tc.setHeaderValue("<HTML><B>Gmail</B></HTML>");

		} else if (selectedInput.equals(InputSource.Bulk.getValue())) {
			TableColumn tc = tcm.getColumn(0);
			tc.setHeaderValue("<HTML><B>Account Name</B></HTML>");
			tc = tcm.getColumn(1);
			tc.setHeaderValue("<HTML><B>Folder Name</B></HTML>");
			tc = tcm.getColumn(2);
			tc.setHeaderValue("<HTML><B>Status</B></HTML>");
			tc = tcm.getColumn(3);
			tc.setHeaderValue("<HTML><B>All Mail Count</B></HTML>");
			tc = tcm.getColumn(4);
			tc.setHeaderValue("<HTML><B>Total Mail In Folder</B></HTML>");
		}

		else {

			TableColumn tc = tcm.getColumn(0);
			tc.setHeaderValue("<HTML><B>Folders</B></HTML>");
			tc = tcm.getColumn(1);
			tc.setHeaderValue("<HTML><B>Folder Name</B></HTML>");
			tc = tcm.getColumn(2);
			tc.setHeaderValue("<HTML><B>Error Count</B></HTML>");
			tc = tcm.getColumn(3);
			tc.setHeaderValue("<HTML><B>All Mail Count</B></HTML>");
			tc = tcm.getColumn(4);
			tc.setHeaderValue("<HTML><B>Total Mail In Folder</B></HTML>");

		}

		th.repaint();
	}

	public void InputGmailAPPLogin() {

		if (!InputtxtUserName.getText().isEmpty()) {

			try {
				LogUtils.setTextToLogScreen(textPane_log, logger, "Inside input Gmail APP Login Process");
				input_password = new String(passwordField_1.getPassword()).trim();
				input_userName = new String(InputtxtUserName.getText()).trim();

				googleLogin = new GoogleLogin();
				googleLogin.googleOathCredentials(input_userName,SELCTION_TYPE_INPUT);
				inputGmailCredential=googleLogin.getoathGoogleCredential();

			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				System.out.println(e.getMessage());
			}

			buttonEnables();
		} else {

			logger.warn("field cannot be empty!!!");
			JOptionPane.showMessageDialog(EmailWizardApplication.this, "field cannot be empty!!!",
					ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

			buttonEnables();
		}
	}

	public void OutputGmailAppLogin() {

		if (!outputUsernameField.getText().isEmpty()) {

			try {
				LogUtils.setTextToLogScreen(textPane_log, logger, "Inside output Gmail APP Login Process");
				output_userName = new String(outputUsernameField.getText()).trim();				
				googleLogin = new GoogleLogin();
				googleLogin.googleOathCredentials(output_userName,SELCTION_TYPE_OUTPUT);
				outputGmailCredential= googleLogin.getoathGoogleCredential();
				LogUtils.setTextToLogScreen(textPane_log, logger,"Connection done with : " + output_userName);

			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				System.out.println(e.getMessage());
			}

			buttonEnables();
		} else {

			logger.warn("field cannot be empty!!!");
			JOptionPane.showMessageDialog(EmailWizardApplication.this, "field cannot be empty!!!",
					ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

			buttonEnables();
		}
	}

	public void GSuiteLogin() {

		if (!txtServiceAccountIDorImapHostName.getText().isEmpty() && !InputtxtUserName.getText().isEmpty()
				&& !textField_p12FileAndPortNo.getText().isEmpty()) {

			try {
				LogUtils.setTextToLogScreen(textPane_log, logger, "Inside G-Suite Login Process");
				String serviceAccountID = txtServiceAccountIDorImapHostName.getText().trim();
				String userName = InputtxtUserName.getText().trim();
				String p12File = textField_p12FileAndPortNo.getText().trim();

				googleLogin = new GoogleLogin();
				googleLogin.googleCredentials(serviceAccountID, userName, p12File);

			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				System.out.println(e.getMessage());
			}

			buttonEnables();
		} else {

			logger.warn("field cannot be empty!!!");

			JOptionPane.showMessageDialog(EmailWizardApplication.this, "field cannot be empty!!!",
					ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

			buttonEnables();
		}
	}

	public boolean checkImapValidations() {
		int passwordLength = passwordField_1.getPassword().length;
		if (InputSource.IMAP.getValue().equals(selectedInput) || InputSource.HOSTGATOR.getValue().equals(selectedInput)
				|| chckbx_Proxy.isSelected()) {
			if ((!InputtxtUserName.getText().isEmpty()) && passwordLength > 0
					&& !txtServiceAccountIDorImapHostName.getText().isEmpty()
					&& !textField_p12FileAndPortNo.getText().isEmpty()) {
				return true;
			}
		} else {

			if (!InputtxtUserName.getText().isEmpty() && passwordLength > 0) {
				return true;
			}
		}

		return false;

	}

	public void IMAPLogin() {

		if (checkImapValidations()) {

			try {
				LogUtils.setTextToLogScreen(textPane_log, logger, "Connecting with : " + InputtxtUserName.getText());
				count = 0;
				input_userName = new String(InputtxtUserName.getText()).trim();
				input_password = new String(passwordField_1.getPassword()).trim();

				clientforimap_input = connectionWithInputIMAP();
				clientforimap_input.setUseMultiConnection(MultiConnectionMode.Enable);

				LogUtils.setTextToLogScreen(textPane_log, logger, "Connection done with : " + InputtxtUserName.getText());

				CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
				card.show(EmailWizardApplication.CardLayout, "GoogleUserDetailsPanel_2");

				ImapFolderInfoCollection folderInfoColl = clientforimap_input.listFolders();
				model = (DefaultTableModel) EmailWizardApplication.table_UserDetails.getModel();
				getImapFolders(folderInfoColl);

			} catch (Exception ex) {
				btnEnabled();
				logger.error("Error Found ", ex);
				ExceptionHandler exceptionHandler = new ExceptionHandler(ex, frame);
				exceptionHandler.loginExceptionHandler();
			}

			buttonEnables();
		} else {

			logger.warn("field cannot be empty!!!");

			JOptionPane.showMessageDialog(EmailWizardApplication.this, "field cannot be empty!!!",
					ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

			buttonEnables();
		}
	}

	public void EWSLogin() {

		if (!InputtxtUserName.getText().isEmpty()) {

			try {
				LogUtils.setTextToLogScreen(textPane_log, logger, "Trying to connect with EWS");
				count = 0;
				input_password = new String(passwordField_1.getPassword()).trim();
				input_userName = new String(InputtxtUserName.getText()).trim();
				ews = new EWSOffice();
				// service = ews.loginEWS(input_userName, input_password);
				service = ews.loginEWS(input_userName);

				service.findFolders(WellKnownFolderName.Root, new FolderView(1));
				LogUtils.setTextToLogScreen(textPane_log, logger, "Connection done with EWS Server");

				CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
				card.show(EmailWizardApplication.CardLayout, "p_msoffice");

			} catch (Exception ex) {
				btnEnabled();
				ex.printStackTrace();
				ExceptionHandler exceptionHandler = new ExceptionHandler(ex, frame);
				exceptionHandler.loginExceptionHandler();
			}

			buttonEnables();
		} else {
			LogUtils.setTextToLogScreen(textPane_log, logger, "Fields cannot be empty!!!");
			JOptionPane.showMessageDialog(EmailWizardApplication.this, "field cannot be empty!!!",
					ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));

			buttonEnables();
		}

	}

	public static ExchangeService resetEWS() {
		return ews.loginEWS(input_userName, input_password);
	}

	public void btnDisable() {
		lblimapGif.setVisible(true);
		SavingOptionPanel.setEnabled(false);
		btn_login.setEnabled(false);
	}

	public void btnEnabled() {
		lblimapGif.setVisible(false);
		SavingOptionPanel.setEnabled(true);
		btn_login.setEnabled(true);
	}

	public void fileSavingFormatEvent() {
		chckbxNamingconvention.setVisible(true);
		comboBoxNamingConvention.setVisible(true);
		checkBoxSplitPst.setVisible(false);
		radioButtonMB.setVisible(false);
		rdbtnGb.setVisible(false);
		spinner_GB.setVisible(false);
		spinner_MB.setVisible(false);
		chckbxSaveSeperateAttachments.setVisible(true);
	}

	public void radioButtionEvent(ItemEvent e) {
		
		if (e.getStateChange() == ItemEvent.SELECTED) {

			outputUsernameField.setVisible(true);
			passwordField.setVisible(true);
			lblNewLabel_Useranme.setVisible(true);
			lblUsername_1.setVisible(true);
			lblNewLabel_Password.setVisible(true);
			btn_login.setVisible(true);
			btnNew_p3.setVisible(false);
			textField_DownloadingPath.setVisible(false);
			btnDownloadingPath.setVisible(false);
			if (r_imap.isSelected() || r_hostgator.isSelected()) {
				textField_portOutput.setVisible(true);
				txtCloudhostgatorcom.setVisible(true);
				lblNewLabel.setVisible(true);
				lbl_Hostoutput.setVisible(true);
			} else if (r_gmail_app.isSelected()) {

				passwordField.setText("gmailapp");
				passwordField.setVisible(false);
				lblNewLabel_Password.setVisible(false);
			}

		} else if (e.getStateChange() == ItemEvent.DESELECTED) {
			outputUsernameField.setVisible(false);
			passwordField.setVisible(false);
			lblNewLabel_Useranme.setVisible(false);
			lblNewLabel_Password.setVisible(false);
			lblUsername_1.setVisible(false);
			btn_login.setVisible(false);
			btnNew_p3.setVisible(true);
			textField_DownloadingPath.setVisible(true);
			btnDownloadingPath.setVisible(true);
			textField_portOutput.setVisible(false);
			txtCloudhostgatorcom.setVisible(false);
			lblNewLabel.setVisible(false);
			lbl_Hostoutput.setVisible(false);

		}
     boolean ischeckEmailClientSelected = OutputSource.imapClientOutputFormat.contains(emailClientSelectedFormatAtOutput());		
		if(ischeckEmailClientSelected)
		{
			l_contact.setVisible(false);
			l_calendar.setVisible(false);
			l_drive.setVisible(false);
			l_photos.setVisible(false);
			
			c_calendar.setVisible(false);
			c_contact.setVisible(false);			
			c_drive.setVisible(false);			
			c_photos.setVisible(false);	
			
			c_photos.setSelected(false);
			c_contact.setSelected(false);
			c_drive.setSelected(false);
			c_calendar.setSelected(false);
		
			chckbxSaveSeperateAttachments.setVisible(false);
			chckbxSaveSeperateAttachments.setSelected(false);
		}
		else
		{
			l_contact.setVisible(true);
			l_calendar.setVisible(true);
			l_drive.setVisible(true);
			l_photos.setVisible(true);
			c_calendar.setVisible(true);
			c_contact.setVisible(true);
			c_drive.setVisible(true);
			c_photos.setVisible(true);
			c_photos.setSelected(false);
			c_contact.setSelected(false);
			c_drive.setSelected(false);
			c_calendar.setSelected(false);
			chckbxSaveSeperateAttachments.setVisible(true);
		}
	}

	static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
				.replace("//s", "").replace("\"", "");
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName;
	}

	public void TextFieldPopup(JTextField textField) {

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

		textField.setComponentPopupMenu(menu);
	}

	private static HttpRequestInitializer setHttpTimeout(final HttpRequestInitializer requestInitializer) {
		return new HttpRequestInitializer() {
			@Override
			public void initialize(HttpRequest httpRequest) throws IOException {
				requestInitializer.initialize(httpRequest);
				httpRequest.setConnectTimeout(3 * 60000); // 3 minutes connect timeout
				httpRequest.setReadTimeout(3 * 60000); // 3 minutes read timeout
			}

		};
	}
	public Gmail getoutputGmailAppService() throws GeneralSecurityException, IOException
	{
		NetHttpTransport  HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		 outputGmailService = new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(outputGmailCredential)).setApplicationName(APPLICATION_NAME).build();
		 return outputGmailService;		
	}

	class HeaderRenderer extends JCheckBox implements TableCellRenderer {
		public HeaderRenderer(JTableHeader header, final int targetColumnIndex) {
			super((String) null);
			setOpaque(false);
			setFont(header.getFont());
			header.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseClicked(MouseEvent e) {
					JTableHeader header = (JTableHeader) e.getSource();
					JTable table = header.getTable();
					TableColumnModel columnModel = table.getColumnModel();
					int vci = columnModel.getColumnIndexAtX(e.getX());
					int mci = table.convertColumnIndexToModel(vci);
					if (mci == targetColumnIndex) {
						TableColumn column = columnModel.getColumn(vci);
						Object v = column.getHeaderValue();
						boolean b = Status.DESELECTED.equals(v) ? true : false;
						TableModel m = table.getModel();
						for (int i = 0; i < m.getRowCount(); i++)
							m.setValueAt(b, i, mci);
						column.setHeaderValue(b ? Status.SELECTED : Status.DESELECTED);
					}
				}
			});
		}

		@Override
		public Component getTableCellRendererComponent(JTable tbl, Object val, boolean isS, boolean hasF, int row,
				int col) {
			if (val instanceof Status) {
				switch ((Status) val) {
				case SELECTED:
					setSelected(true);
					setEnabled(true);
					break;
				case DESELECTED:
					setSelected(false);
					setEnabled(true);
					break;
				case INDETERMINATE:
					setSelected(true);
					setEnabled(false);
					break;
				}
			} else {
				setSelected(true);
				setEnabled(false);
			}
			TableCellRenderer r = tbl.getTableHeader().getDefaultRenderer();
			JLabel l = (JLabel) r.getTableCellRendererComponent(tbl, null, isS, hasF, row, col);

			l.setIcon(new CheckBoxIcon(this));
			l.setText(null);
			l.setHorizontalAlignment(SwingConstants.CENTER);

			return l;
		}
	}

	class HeaderCheckBoxHandler implements TableModelListener {
		private final JTable table;

		public HeaderCheckBoxHandler(JTable table) {
			this.table = table;
		}

		@Override
		public void tableChanged(TableModelEvent e) {
			if (e.getType() == TableModelEvent.UPDATE && e.getColumn() == 3) {
				int mci = 3;
				int vci = table.convertColumnIndexToView(mci);
				TableColumn column = table.getColumnModel().getColumn(vci);
				Object title = column.getHeaderValue();
				if (!Status.INDETERMINATE.equals(title)) {
					column.setHeaderValue(Status.INDETERMINATE);
				} else {
					int selected = 0, deselected = 0;
					TableModel m = table.getModel();
					for (int i = 0; i < m.getRowCount(); i++) {
						if (Boolean.TRUE.equals(m.getValueAt(i, mci))) {
							selected++;
						} else {
							deselected++;
						}
					}
					if (selected == 0) {
						column.setHeaderValue(Status.DESELECTED);
					} else if (deselected == 0) {
						column.setHeaderValue(Status.SELECTED);
					} else {
						return;
					}
				}
				table.getTableHeader().repaint();
			}
		}
	}

	enum Status {
		SELECTED, DESELECTED, INDETERMINATE
	}

	class CheckBoxIcon implements Icon {
		private final JCheckBox check;

		public CheckBoxIcon(JCheckBox check) {
			this.check = check;
		}

		@Override
		public int getIconWidth() {
			return check.getPreferredSize().width;
		}

		@Override
		public int getIconHeight() {
			return check.getPreferredSize().height;
		}

		@Override
		public void paintIcon(Component c, Graphics g, int x, int y) {
			SwingUtilities.paintComponent(g, check, (Container) c, x, y, getIconWidth(), getIconHeight());
		}
	}

	class LeftAlignHeaderRenderer implements TableCellRenderer {
		@Override
		public Component getTableCellRendererComponent(JTable t, Object v, boolean isS, boolean hasF, int row,
				int col) {
			TableCellRenderer r = t.getTableHeader().getDefaultRenderer();
			JLabel l = (JLabel) r.getTableCellRendererComponent(t, v, isS, hasF, row, col);
			l.setHorizontalAlignment(SwingConstants.LEFT);
			return l;
		}
	}

	public class CellRenderer extends DefaultTableCellRenderer {

		private static final long serialVersionUID = 1L;

		@Override
		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			setToolTipText("<HTML><B>" + value.toString() + "</B></HTML>");

			return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
		}

	}

	class ColumnColorRenderer extends DefaultTableCellRenderer {
		Color backgroundColor, foregroundColor;

		public ColumnColorRenderer(Color backgroundColor, Color foregroundColor) {
			super();
			this.backgroundColor = backgroundColor;
			this.foregroundColor = foregroundColor;
		}

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			cell.setBackground(backgroundColor);
			cell.setForeground(foregroundColor);
			return cell;
		}
	}

	public class PaintTableCellRenderer extends DefaultTableCellRenderer {

		Color foregroundColor;

		public PaintTableCellRenderer(Color foregroundColor) {
			super();
			this.foregroundColor = foregroundColor;
		}

		@Override
		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			super.getTableCellRendererComponent(table, "", isSelected, hasFocus, row, column);
			if (value instanceof Double) {
				double distance = (double) value;
				int part = (int) (255 * distance);
				Color color = new Color(part, part, part);
				setBackground(color);
			} else {
				setBackground(foregroundColor);
			}
			return this;
		}

	}
}
