package com.download.email;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.security.GeneralSecurityException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.List;
import java.util.NoSuchElementException;

import org.slf4j.LoggerFactory;

import com.api.google.GoogleLogin;
import com.aspose.cells.Workbook;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.IConnection;
import com.aspose.email.ImapClient;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiMessage;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.PersonalStorage;
import com.aspose.word.DownloadWordFormat;
import com.constants.InputSource;
import com.exceptions.ExceptionHandler;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.Base64;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.Label;
import com.google.api.services.gmail.model.ListLabelsResponse;
import com.google.api.services.gmail.model.ListMessagesResponse;
import com.google.api.services.gmail.model.Message;
import com.google.api.services.gmail.model.MessagePartHeader;
import com.main.EmailWizardApplication;
import com.util.CSVUtils;
import com.util.FileNamingUtils;
import com.util.LogUtils;
import com.util.RemoveDuplicateEmails;

public class GmailBackup implements FileNamingUtils,CSVUtils{

	int cellNo;
	Workbook workbook;
	private BigDecimal sizeOfEmails;
	private int splitCount;
	private int totalEmails = 1;
	private PersonalStorage pst;
	private FolderInfo folderInfo;
	private File pstFile;
	private MboxrdStorageWriter mbox;
	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";


	private String user = "me";

	private File folderNamePST;
	List<String> duplicateEmailsList;
	@SuppressWarnings("deprecation")
	
	static GoogleCredential gSuiteCredentials;
	Credential inputGmailAPPCredential;
	
	private String p12File;
	private String serviceAccountUser;
	private String serviceAccountId;

	ImapClient clientforimap_Output;
	IConnection iconnforimap_Output;
	String parentImapFolderPath;
	String newimapFolderPath;

	String detinationPath;

	NetHttpTransport HTTP_TRANSPORT;
	File folderName;
	private Label parentLabel;
	private Gmail outputGmailService;
	public static Gmail inputGmailService = null;
	String userName;

	public static org.slf4j.Logger logger = LoggerFactory.getLogger(EmailWizardApplication.class);
	

	public GmailBackup(Credential inputGmailAPPCredential,String serviceAccountUser, String serviceAccountId, String p12File, String detinationPath)throws IOException, GeneralSecurityException {
			
		this.serviceAccountUser = serviceAccountUser;
		this.serviceAccountId = serviceAccountId;
		this.p12File = p12File;
		this.detinationPath = detinationPath;
		this.inputGmailAPPCredential=inputGmailAPPCredential;
		
		folderNamePST = new File(detinationPath);
		sizeOfEmails = new BigDecimal(0);
	
		getGSuiteAndGmailService();
	
		folderName = new File(detinationPath);
		
		createPSt(serviceAccountUser);

	}
	public GmailBackup(Credential inputGmailAPPCredential,String serviceAccountUser, String serviceAccountId, String p12File, String detinationPath,Workbook workbook)
			throws IOException, GeneralSecurityException {
		this.serviceAccountUser = serviceAccountUser;
		this.serviceAccountId = serviceAccountId;
		this.p12File = p12File;
		this.detinationPath = detinationPath;
		this.inputGmailAPPCredential=inputGmailAPPCredential;

		getGSuiteAndGmailService();	
		this.workbook=workbook;

	}

	public GmailBackup(Credential inputGmailAPPCredential,String serviceAccountUser, String serviceAccountId, String p12File,
			ImapClient clientforimap_Output, IConnection iconnforimap_Output, String parentImapFolderPath)
			throws IOException, GeneralSecurityException {
		this.serviceAccountUser = serviceAccountUser;
		this.serviceAccountId = serviceAccountId;
		this.p12File = p12File;
		this.clientforimap_Output = clientforimap_Output;
		this.iconnforimap_Output = iconnforimap_Output;
		this.parentImapFolderPath = parentImapFolderPath;
		this.newimapFolderPath = null;
		this.inputGmailAPPCredential=inputGmailAPPCredential;

		getGSuiteAndGmailService();	

	}
	
	public GmailBackup(Credential inputGmailAPPCredential,String serviceAccountUser, String serviceAccountId, String p12File,
			String userName,Label parentLabel, Gmail outputGmailService, String parentImapFolderPath)
			throws IOException, GeneralSecurityException {
		
		this.serviceAccountUser = serviceAccountUser;
		this.serviceAccountId = serviceAccountId;
		this.p12File = p12File;		
		this.parentImapFolderPath = parentImapFolderPath;
		this.newimapFolderPath = null;
		this.inputGmailAPPCredential=inputGmailAPPCredential;
		this.outputGmailService = outputGmailService;
		this.parentLabel = parentLabel;		
		//this.userName=userName;
		this.userName="me";

		getGSuiteAndGmailService();	

	}

	@SuppressWarnings("deprecation")
	
	public void download()throws IOException {
		
		try {
				
			String folderPathname = null;
			duplicateEmailsList = new ArrayList<String>();
			ListLabelsResponse listResponse = inputGmailService.users().labels().list(user).execute();			
			List<Label> labels = listResponse.getLabels();
			
			if (!labels.isEmpty()) {
				
				for (int i = 0; i < labels.size(); i++) {
					
					if(EmailWizardApplication.stop)
					{
						break;
					}
				
					try {
					LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Downloading : "+ labels.get(i).getName());
			         if (EmailWizardApplication.stop) {break;}
					
					  Label label = labels.get(i);
					
					  ListMessagesResponse messagesResponse = inputGmailService.users().messages().list(user).setLabelIds(Arrays.asList(label.getId())).execute();	
					
						folderPathname=	createSelectedFormatObject(label,detinationPath);
						downloadInPages(messagesResponse, folderPathname, label);
												
					} catch (Exception exception) {
						logger.error(exception.getMessage());
						ExceptionHandler exceptionHandler=new ExceptionHandler(exception);						
						if(exceptionHandler.GsuiteExceptionHandler())
						{
							EmailWizardApplication.lblNoInternetConnection.setVisible(true);
							while (!checkInternet()) {}
							EmailWizardApplication.lblNoInternetConnection.setVisible(false);
							getGSuiteAndGmailService();														
							i--;
						}
					
						}
					finally {
						if(EmailWizardApplication.r_csv.isSelected())
						{
							saveCSV(new File(folderPathname));
							workbook.dispose();
						}						
						
					}
					

				}
			}
		} catch (Exception exception) {
			logger.error(exception.getMessage());
			ExceptionHandler exceptionHandler=new ExceptionHandler(exception);						
			if(exceptionHandler.GsuiteExceptionHandler())
			{
				EmailWizardApplication.lblNoInternetConnection.setVisible(true);
				while (!checkInternet()) {}
				EmailWizardApplication.lblNoInternetConnection.setVisible(false);
				download();
			}		
		} finally {
			if (EmailWizardApplication.r_pst.isSelected()) {
				pst.close();
			}

		}

	}
	
	public String createSelectedFormatObject(Label label,String folderPathname) throws GeneralSecurityException, IOException
	{
		if (EmailWizardApplication.r_pst.isSelected()) {
			String labelBackwordSlash = label.getName().replace("/", "\\");
			folderInfo = pst.getRootFolder().addSubFolder(labelBackwordSlash, true);
			
		} 
		else if (EmailWizardApplication.r_mbox.isSelected()) 
		{
			File folderPath = new File(folderName.getAbsolutePath() + File.separator + label.getName().trim());
			folderPath.mkdirs();
			folderPathname = folderPath.getAbsolutePath();
			String str[] = label.getName().split("/");
			String name = str[str.length - 1];
			mbox = new MboxrdStorageWriter(folderPathname + File.separator + name + ".mbx", false);
			
			
		} 
		else if (EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_gmail.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_yandex.isSelected() ||  EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {
			
			if (EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {
				String removeDotlabelName =label.getName().replace(".", "-");
				String removeSalshWithDotlabelName =removeDotlabelName.replace("/", ".");
				newimapFolderPath = parentImapFolderPath + "." + removeSalshWithDotlabelName;
			} else {
				newimapFolderPath = parentImapFolderPath + "/" + label.getName();
			}
			
				try {
					clientforimap_Output.createFolder(iconnforimap_Output,newimapFolderPath);									
					if(EmailWizardApplication.r_aws.isSelected())
					{
						clientforimap_Output.selectFolder(iconnforimap_Output,newimapFolderPath);
					}
					else
					{
						clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
						clientforimap_Output.subscribeFolder(iconnforimap_Output,newimapFolderPath);
					}
					
				} catch (Error e) {		
					logger.error("ERROR",e);
				} catch (Exception exception) {
					logger.error("ERROR In Imap Migration : " +exception);							
					ExceptionHandler exceptionHandler = new ExceptionHandler(exception);
					if (exceptionHandler.migrationExceptionHandler()) {
						EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to output imap server");
						outputImapServerReconnction();
						clientforimap_Output.createFolder(iconnforimap_Output,newimapFolderPath);										
						if(EmailWizardApplication.r_aws.isSelected())
						{
							clientforimap_Output.selectFolder(iconnforimap_Output,newimapFolderPath);
						}
						else
						{
							clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
							clientforimap_Output.subscribeFolder(iconnforimap_Output,newimapFolderPath);
						}
						
						logger.info("Migration started again");
					}
				
			}
					
		}
		 else if (EmailWizardApplication.r_gmail_app.isSelected()) {
			 
			     newimapFolderPath = parentImapFolderPath + "/" + label.getName();
			    Label newlabel = new Label();
			    newlabel.setName(newimapFolderPath);
			    newlabel.setLabelListVisibility("labelShow");
			    newlabel.setMessageListVisibility("show");
				try {				   
					parentLabel = outputGmailService.users().labels().create(userName, newlabel).execute();
				} catch (Exception exception) {
					logger.error(exception.getMessage());
					ExceptionHandler exceptionHandler=new ExceptionHandler(exception);						
					if(exceptionHandler.GsuiteExceptionHandler())
					{
						EmailWizardApplication.lblNoInternetConnection.setVisible(true);
						while (!checkInternet()) {}
						EmailWizardApplication.lblNoInternetConnection.setVisible(false);
						getOutputGmailAppService();
						parentLabel = outputGmailService.users().labels().create(userName, newlabel).execute();
					}
				}
											

		 }
		else if(EmailWizardApplication.r_csv.isSelected())
		{
			File folderPath = new File(folderPathname + File.separator + label.getName().trim());
			folderPath.mkdirs();
			folderPathname = folderPath.getAbsolutePath();
			workbook=createCSVStructure();
		}
		else {
			File folderPath = new File(folderName.getAbsolutePath() + File.separator + label.getName().trim());
			folderPath.mkdirs();
			folderPathname = folderPath.getAbsolutePath();																			
			
		}
		
		return folderPathname;
		
	}
	
	
	public void downloadInPages(ListMessagesResponse messagesResponse,String folderPathname,Label label) throws IOException, GeneralSecurityException {
		int i=0;
		while (messagesResponse.getMessages() != null) 
		{
			if(i==1){break;}
			List<Message> messages= messagesResponse.getMessages();			
			downloadEachPage(messages, folderPathname, label);
			
			if (messagesResponse.getNextPageToken() != null) {
				String pageToken = messagesResponse.getNextPageToken();
				EmailWizardApplication.downloadingFileName.setText("<html><font color=\"red\">[Please wait Fetching NextPageToken]</font></html>");										
				messagesResponse = inputGmailService.users().messages().list(user).setLabelIds(Arrays.asList(label.getId())).setPageToken(pageToken).execute();
						
			} else {
				break;
			}
			i++;
		}

		
	}
public boolean isDuplicate(MailMessage msg) {
		
		if (EmailWizardApplication.chckbxSkipDuplicate.isSelected()) {
			
			RemoveDuplicateEmails removeDuplicate=RemoveDuplicateEmails.getInstance();
			removeDuplicate.setMailMessage(msg);

			if (duplicateEmailsList.contains(removeDuplicate.getHashString())) {
				return true;
			} else {
										
				duplicateEmailsList.add(removeDuplicate.getHashString());
			}
		}
		return false;
	}
	private void downloadEachPage(List<Message> messages, String folderPath, Label label) throws GeneralSecurityException, IOException {
		
		setProgressBar();
		for (int i = 0; i < messages.size(); i++) {

			try {
				if(checkStopAndDemo(i)){break;}
					
				String id = messages.get(i).getId();
				
				Message msgFull = getMessage(id, "full");
				List<MessagePartHeader> messagePartHeaderList = msgFull.getPayload().getHeaders();
				String subjectName =getsubjectName(messagePartHeaderList,msgFull);
				
				Message msgRaw = getMessage(id, "raw");
				String decodedString = new String(msgRaw.decodeRaw());				
				MailMessage mailMessage= convertByteArrayToMailMessage(decodedString);

				if (EmailWizardApplication.checkBoxSplitPst.isSelected()) {
					kbToMB(msgFull.getSizeEstimate(), label);
				}

				if (EmailWizardApplication.rdbtnDateFilter.isSelected()) {
					if (!checkDateExist(msgFull.getInternalDate())) {
						continue;
					}
				}

				if (isDuplicate(mailMessage)) {continue;}
				
				String destinationPath=folderPath + File.separator + subjectName + "_" + i;
				downloadEmailInselectedFormat(mailMessage,destinationPath);
				
				EmailWizardApplication.downloadingFileName.setText(subjectName + "_" + i);
				EmailWizardApplication.modelDownloading.setValueAt(totalEmails, EmailWizardApplication.rownCount, 4);
				int prog = ((i + 1) * 100) / messages.size();
				EmailWizardApplication.progressBar_Downloading.setValue(prog);
				totalEmails++;

			} catch (Exception exception) {
				logger.error(exception.getMessage());
				ExceptionHandler exceptionHandler=new ExceptionHandler(exception);						
				if(exceptionHandler.GsuiteExceptionHandler())
				{
					EmailWizardApplication.lblNoInternetConnection.setVisible(true);
					while (!checkInternet()) {}
					EmailWizardApplication.lblNoInternetConnection.setVisible(false);
					getGSuiteAndGmailService();								
					i--;
				}
			
				}
			}

		
	}


	@SuppressWarnings("resource")
	public void downloadEmailInselectedFormat(MailMessage msg,String destinationPath)
			throws Exception {

		if (EmailWizardApplication.r_mbox.isSelected()) {

			mbox.writeMessage(msg);

		} else if (EmailWizardApplication.r_pst.isSelected()) {

			folderInfo.addMessage(MapiMessage.fromMailMessage(msg));

		} else if (EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_gmail.isSelected()  || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_yandex.isSelected() ||  EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {
		
				   imapServerRefresh(newimapFolderPath);
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
						getOutputGmailAppService();
						gmailImport(msg);
	
					}
		}
		}
		else if(EmailWizardApplication.r_csv.isSelected())
		{
			saveCSVEmailandTask(msg);
			
		}
		else
		{

			DownloadWordFormat dwf= new DownloadWordFormat();
			dwf.saveWordFormat(msg, destinationPath);
		}

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
	private void appendMessagetoOutPutServer(MailMessage msg) {
		try {
			clientforimap_Output.appendMessage(iconnforimap_Output, msg);
		} catch (Error e) {		
			logger.error("ERROR",e);
		} catch (Exception exception) {
			logger.error("ERROR In Imap Migration : " +exception);
					
			ExceptionHandler exceptionHandler = new ExceptionHandler(exception);
			if (exceptionHandler.migrationExceptionHandler()) {
				EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to output imap server");
				outputImapServerReconnction();
				clientforimap_Output.appendMessage(iconnforimap_Output, msg);		
				logger.info("Migration started again");
			}
		}
	}
	public void outputImapServerReconnction() {

		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Please wait....Trying To Reconnect to output Imap");	
		EmailWizardApplication.lblNoInternetConnection.setVisible(true);
		if (clientforimap_Output != null && iconnforimap_Output != null) {

			clientforimap_Output.close();
			clientforimap_Output.dispose();
			iconnforimap_Output.dispose();
			
		}
		while (true) {
			try {
				clientforimap_Output= EmailWizardApplication.outputImapConnection();
				iconnforimap_Output = clientforimap_Output.createConnection();
                 if(EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {
					
					clientforimap_Output.selectFolder(iconnforimap_Output,newimapFolderPath);
				}
				else
				{
					clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
				}
				
				break;
			} catch (Exception e) {
				
				logger.error("Connection Refresh Error"+e.getMessage());

			}
		}
		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Connected to output Imap server");
		EmailWizardApplication.lblNoInternetConnection.setVisible(false);

	}
	private void imapServerRefresh(String folderinfo)
	{
		if (EmailWizardApplication.checkImapConnectionTime <= System.currentTimeMillis()) {
			
			EmailWizardApplication.lblNoInternetConnection.setText("Imap Connection Refresh");
		
			if(clientforimap_Output!=null&&iconnforimap_Output!=null)
			{
				LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Doing refresh for imap output server");
				outputImapServerReconnction();
				clientforimap_Output.selectFolder(iconnforimap_Output, newimapFolderPath);
			}
			
			EmailWizardApplication.checkImapConnectionTime = System.currentTimeMillis() + EmailWizardApplication.IMAP_RERESH_TIMEOUT;
			LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Imap Connection Refresh done");
																			
		}
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

	public void kbToMB(int size, Label label) {
		boolean checkSize = false;
		BigDecimal kilobytes = BigDecimal.valueOf(pstFile.length()).divide(BigDecimal.valueOf(1024));
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
			pst.close();
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
			LocalDateTime now = LocalDateTime.now();
			System.out.println(dtf.format(now));
			pst = PersonalStorage.create(folderNamePST.getAbsolutePath() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + serviceAccountUser + ".pst", FileFormatVersion.Unicode);					
			pstFile = new File(folderNamePST.getAbsolutePath() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + serviceAccountUser + ".pst");					
			pst.getStore().changeDisplayName(serviceAccountUser);
			folderInfo = new FolderInfo();
			String labelBackwordSlash = label.getName().replace("/", "\\");
			folderInfo = pst.getRootFolder().addSubFolder(labelBackwordSlash, true);
			splitCount++;
		}

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
	public void getGSuiteAndGmailService() throws GeneralSecurityException, IOException
	{
		if (EmailWizardApplication.selectedInput.equals(InputSource.GSUITE.getValue())) {
			getGSuiteService();
		}
		else if (EmailWizardApplication.selectedInput.equals(InputSource.GMAIL_APP.getValue())) {
			getinputGmailAppService();
		}
	}
	public Gmail getGSuiteService() throws GeneralSecurityException, IOException
	{
	  	 HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		   gSuiteCredentials = new GoogleCredential.Builder().setTransport(HTTP_TRANSPORT).setJsonFactory(JSON_FACTORY)
				.setServiceAccountId(serviceAccountId)
				.setServiceAccountScopes(Collections.singleton("https://mail.google.com"))
				.setServiceAccountUser(serviceAccountUser).setServiceAccountPrivateKeyFromP12File(new File(p12File))
				.build();
		        inputGmailService = new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(gSuiteCredentials)).setApplicationName(APPLICATION_NAME).build();
			return inputGmailService;
		
	}
	public Gmail getinputGmailAppService() throws GeneralSecurityException, IOException
	{
	  	 HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		 inputGmailService = new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(inputGmailAPPCredential)).setApplicationName(APPLICATION_NAME).build();
		 return inputGmailService;		
	}
	public Gmail getOutputGmailAppService() throws GeneralSecurityException, IOException
	{
	  	  HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
	  	 GoogleLogin googleLogin = new GoogleLogin();
	  	 outputGmailService = new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(googleLogin.getoathGoogleCredential())).setApplicationName(APPLICATION_NAME).build();
		 return outputGmailService;		
	}
	public Message getMessage(String messageId, String format) throws IOException {

		Gmail.Users.Messages.Get messagesGet = inputGmailService.users().messages().get(user, messageId);
		messagesGet.setFormat(format);
		Message message = messagesGet.execute();
		return message;
	}
	public MailMessage convertByteArrayToMailMessage(String decodedString)
	{
		ByteArrayInputStream bis = new ByteArrayInputStream(decodedString.getBytes());
		return new MailMessage().load(bis);
	}
	
	boolean checkStopAndDemo(int i)
	{
		if (EmailWizardApplication.stop) {
			return true;
		}

		if (EmailWizardApplication.demo) {
			if (i == EmailWizardApplication.DEMO_LIMIT) {
				return true;
			}
		}
		return false;
		
	}
	String getsubjectName(List<MessagePartHeader> messagePartHeaderList,Message msgFull )
	{
		if (EmailWizardApplication.chckbxNamingconvention.isSelected()) {
			return FileNamingUtils.buildFileName(messagePartHeaderList,msgFull.getInternalDate());
		} else {
			try {
				MessagePartHeader lcm = messagePartHeaderList.stream().filter(x -> x.getName().equals("Subject")).findFirst().get();
				return FileNamingUtils.namingConvention(lcm.getValue());
				
			} catch (NoSuchElementException e) {
				// TODO: handle exception
				System.out.println(e.getMessage());
			}

		}
		return "";
		
	}
	public void createPSt(String serviceAccountUser)
	{
		if (EmailWizardApplication.r_pst.isSelected()) {
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
			LocalDateTime now = LocalDateTime.now();
			pst = PersonalStorage.create(folderName.getAbsolutePath() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + serviceAccountUser + ".pst", FileFormatVersion.Unicode);					
			pstFile = new File(folderName.getAbsolutePath() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + serviceAccountUser + ".pst");					
			pst.getStore().changeDisplayName(serviceAccountUser);
			folderInfo = new FolderInfo();
			splitCount++;

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
	    
		outputGmailService.users().messages().insert(userName, message).setInternalDateSource("dateHeader").execute();
	}
	
	public void setProgressBar()
	{
		EmailWizardApplication.progressBar_Downloading.setValue(0);
		EmailWizardApplication.progressBar_Downloading.setVisible(true);
		EmailWizardApplication.progressBar_Downloading.setMaximum(100);
		EmailWizardApplication.progressBar_Downloading.setStringPainted(true);
		EmailWizardApplication.lblDownloading.setVisible(true);
	}
	
	

}