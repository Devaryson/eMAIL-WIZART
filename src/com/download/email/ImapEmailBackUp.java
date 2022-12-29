package com.download.email;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;
import java.util.NoSuchElementException;

import org.slf4j.LoggerFactory;

import com.api.google.GoogleLogin;
import com.aspose.cells.Workbook;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.IConnection;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapFolderInfo;
import com.aspose.email.ImapMessageInfo;
import com.aspose.email.ImapMessageInfoCollection;
import com.aspose.email.ImapPageInfo;
import com.aspose.email.ImapQueryBuilder;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MailQuery;
import com.aspose.email.MapiMessage;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.PageSettings;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.aspose.email.TimeoutException;
import com.aspose.word.DownloadWordFormat;
import com.aspose.words.Document;
import com.constants.InputSource;
import com.constants.OutputSource;
import com.exceptions.ExceptionHandler;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.Base64;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.Label;
import com.google.api.services.gmail.model.ListMessagesResponse;
import com.google.api.services.gmail.model.Message;
import com.main.EmailWizardApplication;
import com.util.CSVUtils;
import com.util.FileNamingUtils;
import com.util.LogUtils;
import com.util.MapiMessageUtils;
import com.util.RemoveDuplicateEmails;


public class ImapEmailBackUp  implements FileNamingUtils, MapiMessageUtils, CSVUtils{

	private List<String> duplicateEmailsList;
	private int totalEmails = 1;
	private ImapClient clientforimap_input;
	private IConnection iconnforimap_input;
	private ImapClient clientforimap_Output;
	private IConnection iconnforimap_Output;
	private FolderInfo pstfolderInfo;
	private File path;
	private String value;
	private String imapFolderPath;
	private int mailCount;
	private int failedMailCount;
	private static final int ITEM_PER_PAGE=100; 
	private MboxrdStorageWriter mbox ;
	private Workbook workbook;
	public int cellNo;
	private String selectedItemAtInput;
	int splitCount=1;
	private Label parentLabel;
	private Gmail outputGmailService;
	String userName;
	private String user = "me";
	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	NetHttpTransport HTTP_TRANSPORT;

	
	public static org.slf4j.Logger logger=LoggerFactory.getLogger(EmailWizardApplication.class);

	public ImapEmailBackUp(String selectedItemAtInput,ImapClient clientforimap_input, IConnection iconnforimap_input, File path, String value) {
		this.clientforimap_input = clientforimap_input;
		this.iconnforimap_input = iconnforimap_input;
		this.path = path;
		this.value = value;
		this.selectedItemAtInput=selectedItemAtInput;
	}
	public ImapEmailBackUp(String selectedItemAtInput,ImapClient clientforimap_input, IConnection iconnforimap_input, File path, String value,Workbook workbook) {
		this.clientforimap_input = clientforimap_input;
		this.iconnforimap_input = iconnforimap_input;
		this.path = path;
		this.value = value;
		this.selectedItemAtInput=selectedItemAtInput;
		this.workbook=workbook;
		createCSVStructure();
	}
	
	public ImapEmailBackUp(String selectedItemAtInput,ImapClient clientforimap_input, IConnection iconnforimap_input,MboxrdStorageWriter mbox , File path, String value) {
		this.clientforimap_input = clientforimap_input;
		this.iconnforimap_input = iconnforimap_input;
		this.mbox=mbox;
		this.path = path;
		this.value = value;
		this.selectedItemAtInput=selectedItemAtInput;
	}

	public ImapEmailBackUp(String selectedItemAtInput,ImapClient clientforimap_input, IConnection iconnforimap_input, PersonalStorage pst,FolderInfo pstfolderInfo, String value) {
			
		this.clientforimap_input = clientforimap_input;
		this.iconnforimap_input = iconnforimap_input;
		this.pstfolderInfo = pstfolderInfo;
		this.value = value;
		this.selectedItemAtInput=selectedItemAtInput;
	}
	public ImapEmailBackUp(String selectedItemAtInput,ImapClient clientforimap_input, IConnection iconnforimap_input,ImapClient clientforimap_Output, IConnection iconnforimap_Output, String imapFolderPath,String value) {
		
		this.clientforimap_input = clientforimap_input;
		this.iconnforimap_input = iconnforimap_input;
		this.clientforimap_Output = clientforimap_Output;
		this.iconnforimap_Output = iconnforimap_Output;
		this.imapFolderPath = imapFolderPath;
		this.value = value;
		this.selectedItemAtInput=selectedItemAtInput;
	}
    public ImapEmailBackUp(String selectedItemAtInput,ImapClient clientforimap_input, IConnection iconnforimap_input,String userName,Label parentLabel, Gmail outputGmailService, String imapFolderPath,String value) {
		
		this.clientforimap_input = clientforimap_input;
		this.iconnforimap_input = iconnforimap_input;
		this.parentLabel = parentLabel;
		this.outputGmailService = outputGmailService;
		this.imapFolderPath = imapFolderPath;
		this.value = value;
		this.selectedItemAtInput=selectedItemAtInput;
		this.userName="me";
	}
	
	public void downloadImapEmails() {
		clientforimap_input.selectFolder(iconnforimap_input, value);
		ImapFolderInfo folderinfo = clientforimap_input.getFolderInfo(iconnforimap_input, value);
		int totalMessages = folderinfo.getTotalMessageCount();
		
		EmailWizardApplication.modelDownloading.setValueAt(folderinfo.getName(), EmailWizardApplication.rownCount, 1);
		EmailWizardApplication.modelDownloading.setValueAt(totalMessages, EmailWizardApplication.rownCount, 4);
		System.out.println("adding in row " +EmailWizardApplication.rownCount);

		final int totalPages = (int) Math.ceil(((double) totalMessages / (double) ITEM_PER_PAGE));

		mailCount = 0;
		failedMailCount = 0;
		ImapPageInfo ImapPageInfo = null;
		for (int pageCount = 0; pageCount < totalPages; pageCount++) {
			try {
				if(checkStopAndDemo(pageCount,1)){break;};
				PageSettings pagesettings = new PageSettings();
				pagesettings.setFolderName(folderinfo.getName());
				pagesettings.setConnection(iconnforimap_input);
				if (EmailWizardApplication.rdbtnDateFilter.isSelected()) {
					ImapPageInfo = clientforimap_input.listMessagesByPage(folderinfo.getName(),queryDateRangeMessage(),ITEM_PER_PAGE);
							
				} else {

					ImapPageInfo = clientforimap_input.listMessagesByPage(ITEM_PER_PAGE, pageCount, pagesettings);

				}

				backupMail(folderinfo, ImapPageInfo);
				
			} catch (OutOfMemoryError error) {			
				logger.error("An exception occurred OutOfMemoryError!",error);
			} catch (Exception exception) {
				logger.error("Conection break at page: "+totalPages,exception);
				ExceptionHandler exceptionHandler=new ExceptionHandler(exception);
				if(exceptionHandler.migrationExceptionHandler())
				{
					EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to input imap server");
					inputImapServerReconnection();					
					clientforimap_input.selectFolder(iconnforimap_input, folderinfo.getName());					
					

				}
				pageCount--;
				logger.info("migratin started from page "+pageCount);
		
			} finally {

				System.gc();
			}

			if (ImapPageInfo.getLastPage()) {

				break;
			}

		}

	}

	private void backupMail(ImapFolderInfo folderinfo, ImapPageInfo ImapPageInfo) {

		ImapMessageInfoCollection messages = ImapPageInfo.getItems();

		duplicateEmailsList = new ArrayList<String>();
		EmailWizardApplication.progressBar_Downloading.setValue(0);
		EmailWizardApplication.progressBar_Downloading.setVisible(true);
		EmailWizardApplication.progressBar_Downloading.setMaximum(100);
		EmailWizardApplication.progressBar_Downloading.setStringPainted(true);
		EmailWizardApplication.lblDownloading.setVisible(true);

		for (int messageCount = 0; messageCount < messages.size(); messageCount++) {

			try {
				ImapMessageInfo msgFull = messages.get_Item(messageCount);
				String subjectName = null;				
				inputImapServerRefresh(folderinfo);
				if(checkStopAndDemo(messageCount,EmailWizardApplication.DEMO_LIMIT))
				{
					break;
				}
				
				if (EmailWizardApplication.chckbxNamingconvention.isSelected()) {
					subjectName = FileNamingUtils.buildFileName(msgFull, msgFull.getDate().getTime());
				} else {
					try {
						subjectName = FileNamingUtils.namingConvention(msgFull.getSubject());
					} catch (NoSuchElementException e) {
						// TODO: handle exception
						System.out.println(e.getMessage());
					}

				}
				
				if (EmailWizardApplication.checkBoxSplitPst.isSelected()) {
					pstSplit(folderinfo.getName());
				}

				 MailMessage msg = fetchMail(msgFull);
				
				if (EmailWizardApplication.chckbxSkipDuplicate.isSelected()) {
					
					RemoveDuplicateEmails removeDuplicate=RemoveDuplicateEmails.getInstance();
					removeDuplicate.setMailMessage(msg);
	
					if (duplicateEmailsList.contains(removeDuplicate.getHashString())) {
						continue;
					} else {
												
						duplicateEmailsList.add(removeDuplicate.getHashString());
					}
				}

				emailSavingFormats(msg, subjectName);

				EmailWizardApplication.downloadingFileName.setText(mailCount + "_" + subjectName);
				EmailWizardApplication.modelDownloading.setValueAt(totalEmails, EmailWizardApplication.rownCount, 3);

				int prog = ((messageCount + 1) * 100) / messages.size();
				EmailWizardApplication.progressBar_Downloading.setValue(prog);

				totalEmails++;
				mailCount++;

			} catch (Error exception) {
				System.out.println("Error " + exception);
				failedMailCount++;
				EmailWizardApplication.modelDownloading.setValueAt(failedMailCount, EmailWizardApplication.rownCount, 2);

			} catch (TimeoutException e) {
				failedMailCount++;
				EmailWizardApplication.modelDownloading.setValueAt(failedMailCount, EmailWizardApplication.rownCount, 2);
				clientforimap_input.selectFolder(iconnforimap_input, folderinfo.getName());
			} catch (Exception exception) {
				logger.error("Conection break for Email count : "+mailCount,exception);
				ExceptionHandler exceptionHandler=new ExceptionHandler(exception);
				if(exceptionHandler.migrationExceptionHandler())
				{
					EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to input imap server");
					inputImapServerReconnection();					
					clientforimap_input.selectFolder(iconnforimap_input, folderinfo.getName());
			
					logger.error("Migration started from email count : "+mailCount,exception);
					messageCount--;
					
				}
			}

		}

	}

	@SuppressWarnings("resource")
	private void emailSavingFormats(MailMessage msg, String subjectName) throws Exception, Error {
		
		if (EmailWizardApplication.r_pst.isSelected()) {
		
			pstfolderInfo.addMessage(MapiMessage.fromMailMessage(msg));
			msg.close();

		} 
		else if (EmailWizardApplication.r_csv.isSelected()) {
			
	
			saveCSVEmailandTask(msg);
			
		}
		else if (EmailWizardApplication.r_mbox.isSelected()) {
		
			  mbox.writeMessage(msg);

		}
		
		else if (EmailWizardApplication.r_office.isSelected()||EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				||EmailWizardApplication.r_gmail.isSelected()||EmailWizardApplication.r_yahoo.isSelected()
				||EmailWizardApplication.r_yandex.isSelected()|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected()
				||EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {			
		
				 appendMessagetoOutPutServer(msg);
		}
		else if (EmailWizardApplication.r_gmail_app.isSelected()) {			
			
			try {
				gmailImport(msg);
											
				} catch (Exception exception) {
					logger.error(exception.getMessage());
					ExceptionHandler exceptionHandler=new ExceptionHandler(exception);						
					if(exceptionHandler.GsuiteExceptionHandler())
					{
						EmailWizardApplication.lblNoInternetConnection.setVisible(true);
						while (!checkInternet()) {}
						EmailWizardApplication.lblNoInternetConnection.setVisible(false);
						getGmailAppService();
						gmailImport(msg);
	
					}
		}
		}
		else 
		{		
			String destinationPath=path.getAbsolutePath() + File.separator + mailCount + "_" + subjectName;
			DownloadWordFormat dwf= new DownloadWordFormat();
			dwf.saveWordFormat(msg, destinationPath);		
		
		}
	}
	
	public MailMessage fetchMail(ImapMessageInfo msgFull)
	{
		MailMessage msg=null;
		if (selectedItemAtInput.equals(OutputSource.ICLOUD.name()))
		{
			msg=clientforimap_input.fetchMessage(iconnforimap_input, msgFull.getSequenceNumber(),true);
		}
		else
		{
			 msg = clientforimap_input.fetchMessage(iconnforimap_input, msgFull.getUniqueId());
			if (msg == null) {
				msg = clientforimap_input.fetchMessage(iconnforimap_input, msgFull.getSequenceNumber(), true);

			}
		}
				
			return msg;
	}
	private void inputImapServerRefresh(ImapFolderInfo folderinfo)
	{
		if (EmailWizardApplication.checkImapConnectionTime <= System.currentTimeMillis()) {
			
			EmailWizardApplication.lblNoInternetConnection.setText("Imap Connection Refresh");
			if(clientforimap_input!=null&&iconnforimap_input!=null)
			{
			
				LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Doing refresh for Imap Input Server");
				inputImapServerReconnection();
				clientforimap_input.selectFolder(iconnforimap_input, folderinfo.getName());	
			}
			if(clientforimap_Output!=null&&iconnforimap_Output!=null)
			{
				LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Doing refresh for imap output server");
				outputImapServerReconnction();
			}
			
			EmailWizardApplication.checkImapConnectionTime = System.currentTimeMillis() + EmailWizardApplication.IMAP_RERESH_TIMEOUT;
			LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Refresh done");
																			
		}
	}

	private void gmailImport(MailMessage msg) throws IOException
	{
		// Encode and wrap the MIME message into a gmail message
		
	    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
	    msg.save(buffer);
	    byte[] rawMessageBytes = buffer.toByteArray();
	    String encodedEmail = Base64.encodeBase64URLSafeString(rawMessageBytes);
	    Message message = new Message(); 
	    message.setLabelIds(Arrays.asList(parentLabel.getId()));
	    message.setRaw(encodedEmail);
	    message.setInternalDate(msg.getDate().getTime());	    
		outputGmailService.users().messages().insert(userName, message).setInternalDateSource("dateHeader").execute();
		
	}
	
	private void appendMessagetoOutPutServer(MailMessage msg) {
		try {
			clientforimap_Output.appendMessage(iconnforimap_Output, msg);
		} catch (Error e) {
			System.out.println(e.getMessage());
		} catch (Exception exception) {
			logger.error("Conection break in migration of email to output imap server email count : " + mailCount,
					exception);
			ExceptionHandler exceptionHandler = new ExceptionHandler(exception);
			if (exceptionHandler.migrationExceptionHandler()) {
				EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to input imap server");
				outputImapServerReconnction();
				clientforimap_Output.appendMessage(iconnforimap_Output, msg);
		
				logger.error("migrating email again email count : " + mailCount, exception);

			}
			else if(exceptionHandler.migrationExceptionHandler())
			{
				EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to output imap server");
				outputImapServerReconnction();
			}

		}
	}

	public void inputImapServerReconnection() {

		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Please wait....Trying To Reconnect to input imap server");
		EmailWizardApplication.lblNoInternetConnection.setVisible(true);
		if (clientforimap_input != null && iconnforimap_input != null) {

			clientforimap_input.close();
			clientforimap_input.dispose();
			iconnforimap_input.dispose();
			if (!iconnforimap_input.isDisposed()) {
				clientforimap_input.dispose();
			}

		}
		while (true) {
			try {
				clientforimap_input = EmailWizardApplication.connectionWithInputIMAP();
				iconnforimap_input = clientforimap_input.createConnection();	
				break;
			} catch (Exception e) {

			}
		}
		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Connected to Input Imap server");
		EmailWizardApplication.lblNoInternetConnection.setVisible(false);

	}
	public void outputImapServerReconnction() {

		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Please wait....Trying To Reconnect to output Imap");
		
		EmailWizardApplication.lblNoInternetConnection.setVisible(true);
		if (clientforimap_Output != null && iconnforimap_Output != null) {

			clientforimap_Output.close();
			clientforimap_Output.dispose();
			iconnforimap_Output.dispose();
			if (!iconnforimap_Output.isDisposed()) {
				clientforimap_Output.dispose();
			}
		}
		while (true) {
			try {
				clientforimap_Output= EmailWizardApplication.outputImapConnection();
				iconnforimap_Output = clientforimap_Output.createConnection();
                 if(EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {
					
					clientforimap_Output.selectFolder(iconnforimap_Output,"INBOX." +  imapFolderPath);
				}
				else
				{
					clientforimap_Output.selectFolder(iconnforimap_Output, imapFolderPath);
				}
				
				break;
			} catch (Exception e) {

			}
		}
		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Connected to output Imap server");
		EmailWizardApplication.lblNoInternetConnection.setVisible(false);

	}
	

	public MailQuery queryDateRangeMessage() {

		Calendar calendarstartdate = EmailWizardApplication.start_dateChooser.getCalendar();
		calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
		calendarstartdate.set(Calendar.MINUTE, 00);
		calendarstartdate.set(Calendar.SECOND, 00);

		Calendar calendarenddate = EmailWizardApplication.end_dateChooser.getCalendar();
		calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
		calendarenddate.set(Calendar.MINUTE, 59);
		calendarenddate.set(Calendar.SECOND, 59);

		ImapQueryBuilder builder = new ImapQueryBuilder();
		builder.getInternalDate().beforeOrEqual(calendarenddate.getTime());
		builder.getInternalDate().since(calendarstartdate.getTime());
		return builder.getQuery();

	}

	public boolean checkDateExist(Long emailMilliseconds) {
		Calendar calendarstartdate = EmailWizardApplication.start_dateChooser.getCalendar();
		calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
		calendarstartdate.set(Calendar.MINUTE, 00);
		calendarstartdate.set(Calendar.SECOND, 00);

		Calendar calendarenddate = EmailWizardApplication.end_dateChooser.getCalendar();
		calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
		calendarenddate.set(Calendar.MINUTE, 59);
		calendarenddate.set(Calendar.SECOND, 59);

		Long startDateMillisecond = calendarstartdate.getTimeInMillis();
		Long endateMillisecond = calendarenddate.getTimeInMillis();

		if (emailMilliseconds >= startDateMillisecond && emailMilliseconds <= endateMillisecond) {
			return true;
		}
		return false;

	}

	
	public Workbook createCSVStructure() {
		cellNo = 1;
		workbook = CSVUtils.createCSVStructure(cellNo);
		cellNo++;
		return workbook;
	}
	public void saveCSVEmailandTask(MailMessage msg) {
		workbook = CSVUtils.saveCSVEmailandTask(workbook, cellNo, msg);
		cellNo++;
	}
	public Workbook saveCSV(File finalDestination) {
		return CSVUtils.saveCSV(workbook,finalDestination);
		
	}
	boolean checkStopAndDemo(int expectedCount,int demoCount)
	{
		if (EmailWizardApplication.demo) {
			if (expectedCount == demoCount) {
				return true;
			}
		}
		if(EmailWizardApplication.stop)
		{
			return true;
		}

		return false;
	}
	public void pstSplit(String folderPath) {
		boolean checkSize = false;
		BigDecimal kilobytes = BigDecimal.valueOf(EmailWizardApplication.pstSplitFile.length()).divide(BigDecimal.valueOf(1024));
		BigDecimal megabytes = kilobytes.divide(BigDecimal.valueOf(1024));
		BigDecimal gigabytes = megabytes.divide(BigDecimal.valueOf(1024));
		BigDecimal value = null;
		if (EmailWizardApplication.rdbtnGb.isSelected()) {
			value = BigDecimal.valueOf((Integer) EmailWizardApplication.spinner_GB.getValue());

			if (gigabytes.compareTo(value) == 1) {
				System.out.println("Size exceed----- " + megabytes + "GB");
				checkSize = true;
			}
		} else if (EmailWizardApplication.radioButtonMB.isSelected()) {
			value = BigDecimal.valueOf((Integer) EmailWizardApplication.spinner_MB.getValue());
			if (megabytes.compareTo(value) == 1) {
				System.out.println("Size exceed----- " + megabytes + "MB");
				checkSize = true;
			}
		}

		if (checkSize) {
			String pstFileName=EmailWizardApplication.pst.getStore().getDisplayName();
			EmailWizardApplication.pst.close();
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
			LocalDateTime now = LocalDateTime.now();			
			String s=EmailWizardApplication.pstSplitFile.getParent();
			EmailWizardApplication.pst = PersonalStorage.create(EmailWizardApplication.pstSplitFile.getParent() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + pstFileName + ".pst", FileFormatVersion.Unicode);					
			EmailWizardApplication.pstSplitFile = new File(EmailWizardApplication.pstSplitFile.getParent() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + pstFileName + ".pst");					
			EmailWizardApplication.pst.getStore().changeDisplayName(pstFileName);
			pstfolderInfo = new FolderInfo();
			String labelBackwordSlash = validateImapFolderName(folderPath).replace("/", "\\");
			pstfolderInfo = EmailWizardApplication.pst.getRootFolder().addSubFolder(labelBackwordSlash, true);
			splitCount++;
		}

	}
	public String validateImapFolderName(String folderPath)
	{		
		if (EmailWizardApplication.selectedInput.equals(InputSource.IMAP.getValue())||EmailWizardApplication.selectedInput.equals(InputSource.HOSTGATOR.getValue())) {
			folderPath=folderPath.replace("INBOX.", "").replace(".", "/");
			
		}
		return folderPath;
	}
	public Gmail getGmailAppService() throws GeneralSecurityException, IOException
	{
	  	 HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
	  	GoogleLogin	googleLogin = new GoogleLogin();
	  	outputGmailService = new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(googleLogin.getoathGoogleCredential())).setApplicationName(APPLICATION_NAME).build();
		 return outputGmailService;		
	}
	public static boolean checkInternet() {
		try {
			URL url = new URL("http://www.google.com");
			URLConnection connection = url.openConnection();
			connection.connect();
			System.out.println("Internet is connected");
			return true;
		} catch (MalformedURLException e) {

		} catch (IOException e) {

		}
		return false;
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
	

}
