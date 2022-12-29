package com.download.email;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
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
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;

import org.slf4j.LoggerFactory;

import com.api.google.GoogleLogin;
import com.aspose.cells.Workbook;
import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Attachment;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.IConnection;
import com.aspose.email.ImapClient;
import com.aspose.email.MailAddress;
import com.aspose.email.MailAddressCollection;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiNote;
import com.aspose.email.MapiTask;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.NoteColor;
import com.aspose.email.NoteSaveFormat;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.aspose.pdf.internal.imaging.coreexceptions.ImageException;
import com.aspose.word.DownloadWordFormat;
import com.aspose.words.Document;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.constants.InputSource;
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
import com.google.api.services.gmail.model.Message;
import com.google.common.io.Files;
import com.main.EmailWizardApplication;
import com.util.CSVUtils;
import com.util.FileNamingUtils;
import com.util.LogUtils;
import com.util.MapiMessageUtils;
import com.util.RemoveDuplicateEmails;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.remote.ServiceRequestException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Contact;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.item.Task;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.core.service.schema.ContactSchema;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.core.service.schema.TaskSchema;
import microsoft.exchange.webservices.data.notification.StreamingSubscriptionConnection;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class MsOfficeBackup implements FileNamingUtils, MapiMessageUtils, CSVUtils {

	private List<String> duplicateEmailsList;
	private ExchangeService service;
	private String folderName;
	private File destination;
	private Map<String, Folder> mapKey;
	final static int PAGE_SIZE = 1;
	private int failedCount = 0;
	private int overAllCount = 0;
	private int folderMailCount = 0;
	private PersonalStorage pst;
	private FolderInfo pstfolderInfo;
	private MboxrdStorageWriter mbox;
	private File fileFormat;
	private IConnection iconnforimap_Output;
	public ImapClient clientforimap_Output;
	public String imapFolderPath;
	public int cellNo;
	Workbook workbook;
	int splitCount=1;
	private Label parentLabel;
	private Gmail outputGmailService;
	String userName;
	
	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	NetHttpTransport HTTP_TRANSPORT;

	List<String> downloadFolderIdlist = new ArrayList<>();
	public static org.slf4j.Logger logger = LoggerFactory.getLogger(EmailWizardApplication.class);

	public MsOfficeBackup(ExchangeService service, Map<String, Folder> mapKey, String folderName, File destination)
			throws Exception {

		// ----file format backup

		this.service = service;
		this.folderName = folderName;
		this.destination = destination;
		this.mapKey = mapKey;

		processInsideFolders();
	}

	public MsOfficeBackup(ExchangeService service, Map<String, Folder> mapKey, String folderName, File destination,
			Workbook workbook) throws Exception {

		// ---- CSV format backup

		this.service = service;
		this.folderName = folderName;
		this.destination = destination;
		this.mapKey = mapKey;
		this.workbook = workbook;
		processInsideCSVFolders();
	}

	public MsOfficeBackup(ExchangeService service, Map<String, Folder> mapKey, String folderName, File destination,
			MboxrdStorageWriter mbox) throws Exception {
		// ----MBOX backup

		this.service = service;
		this.folderName = folderName;
		this.mbox = mbox;
		this.mapKey = mapKey;
		this.destination = destination;
		processInsideMBOXFolders();

	}

	public MsOfficeBackup(ExchangeService service, Map<String, Folder> mapKey, String folderName, PersonalStorage pst,
			FolderInfo pstfolderInfo) throws Exception {
		// ----PST backup

		this.service = service;
		this.folderName = folderName;
		this.pst = pst;
		this.pstfolderInfo = pstfolderInfo;
		this.mapKey = mapKey;
		processInsidePSTFolders();
	}

	public MsOfficeBackup(ExchangeService service, Map<String, Folder> mapKey, String folderName, String imapFolderPath,
			ImapClient clientforimap_Output, IConnection iconnforimap_Output) throws Exception {

		// ----Imap Client backup

		this.service = service;
		this.folderName = folderName;
		this.mapKey = mapKey;
		this.imapFolderPath = imapFolderPath;
		this.clientforimap_Output = clientforimap_Output;
		this.iconnforimap_Output = iconnforimap_Output;
		processInsideImapClientFolders();

	}
	public MsOfficeBackup(ExchangeService service, Map<String, Folder> mapKey, String folderName, String imapFolderPath,
			String userName,Label parentLabel, Gmail outputGmailService) throws Exception {

		// ----Gmail App Client backup

		this.service = service;
		this.folderName = folderName;
		this.mapKey = mapKey;
		this.imapFolderPath = imapFolderPath;
		this.outputGmailService = outputGmailService;
		this.parentLabel = parentLabel;		
		//this.userName=userName;
		this.userName="me";
		processInsideGmailAPPFolders();

	}

	public void processInsideFolders() throws Exception {
		if (mapKey.containsKey(folderName)) {
			Folder folder = mapKey.get(folderName);
			downloadFolder(folder, destination); // download parent folder
			downloadSubFolder(service, folder.getId(), destination);// download sub-folder
			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 4);
		}
	}

	public void processInsideCSVFolders() throws Exception {
		if (mapKey.containsKey(folderName)) {
			Folder folder = mapKey.get(folderName);
			if (folder.getFolderClass().contains("IPF.Contact")) {
				workbook = createCSVStructureContact();
			} else {
				workbook = createCSVStructure();
			}
			downloadFolder(folder, workbook); // download parent folder
			saveCSV(workbook, destination);
			downloadSubFolderForCSV(service, folder.getId(), destination);// download sub-folder
			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 4);
		}
	}

	public void processInsideMBOXFolders() throws Exception {
		if (mapKey.containsKey(folderName)) {
			Folder folder = mapKey.get(folderName);
			downloadFolder(folder, mbox); // download parent folder
			downloadSubFolderForMBOX(service, folder.getId(), destination);// download sunfolder folder
			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 4);
		}
	}

	public void processInsidePSTFolders() throws Exception {
		if (mapKey.containsKey(folderName)) {
			Folder folder = mapKey.get(folderName);
			downloadFolder(folder, pstfolderInfo); // download parent folder
			downloadSubFolderForPST(service, folder, folder.getDisplayName());// download subfolder folder
			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 4);
		}
	}

	public void processInsideImapClientFolders() throws Exception {
		if (mapKey.containsKey(folderName)) {
			Folder folder = mapKey.get(folderName);
			downloadFolder(folder, imapFolderPath); // download parent folder
			downloadSubFolderForImapClient(service, folder.getId(), imapFolderPath);// download subfolder
			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 4);
		}
	}
	public void processInsideGmailAPPFolders() throws Exception {
		if (mapKey.containsKey(folderName)) {
			Folder folder = mapKey.get(folderName);
			downloadFolder(folder, imapFolderPath); // download parent folder
			downloadSubFolderForGmailApp(service, folder.getId(), imapFolderPath);// download subfolder
			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 4);
		}
	}

	public void downloadFolder(Folder folder, Object value) {

		try {

			setUpProgressBar(folder);
			int offSet = 0;
			ItemView itemView = new ItemView(PAGE_SIZE, offSet);

			for (int i = 0; i < folder.getTotalCount(); i++) {

				try {

					itemView.setPropertySet(new PropertySet(BasePropertySet.IdOnly));
					FindItemsResults<Item> searchResults = getSearchResults(folder, itemView);

					backupFolderData(folder, searchResults, destination, value);

					offSet += PAGE_SIZE;
					itemView = new ItemView(PAGE_SIZE, offSet);

					if (EmailWizardApplication.demo) {
						if (folderMailCount == EmailWizardApplication.DEMO_LIMIT) {
							break;
						}
					}
					if (EmailWizardApplication.stop) {
						break;
					}
					

				} catch (ServiceRequestException serviceRequestException) {

					boolean isInternetdisconnect = checkInternetconnection(serviceRequestException, folder.getId(),
							itemView);

					if (isInternetdisconnect) {

						failedCount++;
						EmailWizardApplication.modelDownloading.setValueAt(failedCount, EmailWizardApplication.rownCount, 2);

						offSet += PAGE_SIZE;
						itemView = new ItemView(PAGE_SIZE, offSet);

					} else {
						i--;
					}

				} catch (ServiceLocalException e) {
					logger.error("An exception occurred!", e);
					failedCount++;
					EmailWizardApplication.modelDownloading.setValueAt(failedCount, EmailWizardApplication.rownCount, 2);

				} catch (NullPointerException exception) {

					exception.printStackTrace();
					failedCount++;
					EmailWizardApplication.modelDownloading.setValueAt(failedCount, EmailWizardApplication.rownCount, 2);
				}

				catch (Exception exception) {
					exception.printStackTrace();

					ExceptionHandler exceptionHandler = new ExceptionHandler(exception);
					if (exceptionHandler.migrationExceptionHandler()) {
						EmailWizardApplication.lblNoInternetConnection.setText("No Internet Connection....Trying To Reconnect to input imap server");
						imapOutputConnectionAgain();
						i--;
					} else if (exceptionHandler.appendExceptionHandler()) {
						failedCount++;
						offSet += PAGE_SIZE;
						itemView = new ItemView(PAGE_SIZE, offSet);
						EmailWizardApplication.modelDownloading.setValueAt(failedCount, EmailWizardApplication.rownCount, 2);
					} else {
						failedCount++;
						offSet += PAGE_SIZE;
						itemView = new ItemView(PAGE_SIZE, offSet);
						EmailWizardApplication.modelDownloading.setValueAt(failedCount, EmailWizardApplication.rownCount, 2);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			logger.error("An exception occurred!", e);
		}
	}

	private void backupFolderData(Folder folder, FindItemsResults<Item> searchResults, File destination, Object value)throws ServiceLocalException, Exception {
			
		for (Item item : searchResults) {

			Item itemType = Item.bind(service, item.getId(), new PropertySet(BasePropertySet.IdOnly));
			
			if (EmailWizardApplication.checkBoxSplitPst.isSelected()) {
				pstSplit(folder.getDisplayName());
				value=(FolderInfo)EmailWizardApplication.pstfolderInfo;
			}

			if (itemType instanceof EmailMessage) {

				EmailMessage emailMsg=fetchEmailMessage(itemType);
				MailMessage msg=  fetchMailMessage(emailMsg);
				if (isDuplicate(msg)) {continue;}
				downloadEmailMessage(emailMsg,msg,folderMailCount, value);

			} else if (itemType instanceof Appointment) {

				downloadAppointment(itemType, folderMailCount, value);

			} else if (itemType instanceof Contact) {

				downloadContact(itemType, folderMailCount, value);

			} else if (itemType instanceof Task) {

				downloadTask(itemType, folderMailCount, value);

			}

			EmailWizardApplication.modelDownloading.setValueAt(overAllCount, EmailWizardApplication.rownCount, 3);
			EmailWizardApplication.modelDownloading.setValueAt(folder.getDisplayName() + "/" + folderMailCount,EmailWizardApplication.rownCount, 1);
					

			int prog = ((folderMailCount + 1) * 100) / folder.getTotalCount();
			EmailWizardApplication.progressBar_Downloading.setValue(prog);

			folderMailCount++;
			overAllCount++;

		}
	}
	
	public EmailMessage fetchEmailMessage(Item itemType) throws Exception
	{
		StreamingSubscriptionConnection conn = new StreamingSubscriptionConnection(service, 30);
		PropertySet propertySet = new PropertySet(EmailMessageSchema.MimeContent, EmailMessageSchema.ItemClass);
		return EmailMessage.bind(service, itemType.getId(), propertySet);
	}
	
	public MailMessage fetchMailMessage(EmailMessage emailMsg) throws ServiceLocalException
	{
		ByteArrayInputStream bis = new ByteArrayInputStream(emailMsg.getMimeContent().getContent());
		MailMessage msg = new MailMessage().load(bis);
		return msg;
	}

	private void downloadEmailMessage(EmailMessage emailMsg,MailMessage msg, int mailCount, Object value) throws Exception {

		String subjectName = isNamingConventionSelected(msg);

		if (emailMsg.getItemClass().equals("IPM.StickyNote")) {
			downloadStickyNote(msg, subjectName, mailCount, value);

		} else if (emailMsg.getItemClass().equals("IPM.Activity")) {

			downloadJournal(msg, subjectName, mailCount, value);
		} else {
			downloadEmails(msg, subjectName, mailCount, value);
		}

		EmailWizardApplication.downloadingFileName.setText(folderMailCount + "_" + subjectName);

	}

	@SuppressWarnings("resource")
	private void downloadJournal(MailMessage msg, String subjectName, int mailCount, Object value) throws GeneralSecurityException, IOException {

		if (value instanceof FolderInfo) {
			pstfolderInfo = (FolderInfo) value;
		} else if (value instanceof MboxrdStorageWriter) {
			mbox = (MboxrdStorageWriter) value;
		} else if (value instanceof File) {
			destination = (File) value;
		} else if (value instanceof Workbook) {
			workbook = (Workbook) value;
		}

		MapiConversionOptions m = MapiConversionOptions.getUnicodeFormat();
		m.setPreserveEmbeddedMessageFormat(true);
		m.setForcedRtfBodyForAppointment(true);
		m.setPreserveOriginalAddresses(true);
		m.setPreserveOriginalDates(true);
		MapiMessage mapiMessage = MapiMessage.fromMailMessage(msg, m);
		mapiMessage.setMessageClass("IPM.Activity");
		mapiMessage.setDeliveryTime(msg.getDate());


		if (EmailWizardApplication.r_pst.isSelected()) {

			pstfolderInfo.changeContainerClass("IPF.Journal");
			pstfolderInfo.addMapiMessageItem(mapiMessage);
		} else if (EmailWizardApplication.r_mbox.isSelected()) {
			mbox.writeMessage(msg);
		} else if (EmailWizardApplication.r_gmail.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected() || EmailWizardApplication.r_yandex.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

			breakImapclientAt4min();
			msg = checkMailSize(msg);
			clientforimap_Output.appendMessage(iconnforimap_Output, msg);

		}
		else if (EmailWizardApplication.r_gmail_app.isSelected()) {
			try {
				msg = checkMailSize(msg);
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

		else if (EmailWizardApplication.r_csv.isSelected()) {

			saveCSVEmailandTask(msg);
			msg.close();

		}

		else {
			mapiMessage.save(destination.getAbsolutePath() + File.separator + mailCount + "_" + subjectName + ".msg");
			mapiMessage.close();
			msg.close();

		}

	}

	@SuppressWarnings("resource")
	private void downloadStickyNote(MailMessage msg, String subjectName, int mailCount, Object value) throws GeneralSecurityException, IOException {

		if (value instanceof FolderInfo) {
			pstfolderInfo = (FolderInfo) value;
		} else if (value instanceof MboxrdStorageWriter) {
			mbox = (MboxrdStorageWriter) value;
		} else if (value instanceof File) {
			destination = (File) value;
		} else if (value instanceof Workbook) {
			workbook = (Workbook) value;
		}

		MapiNote note = new MapiNote();
		note.setSubject(msg.getSubject());
		note.setBody(msg.getBody());
		note.setColor(NoteColor.Yellow);
		note.setHeight(500);
		note.setWidth(500);

		if (EmailWizardApplication.r_pst.isSelected()) {

			pstfolderInfo.changeContainerClass("IPF.StickyNote");
			pstfolderInfo.addMapiMessageItem(note);
		} else if (EmailWizardApplication.r_mbox.isSelected()) {
			mbox.writeMessage(msg);
		} else if (EmailWizardApplication.r_gmail.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected() || EmailWizardApplication.r_yandex.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

			breakImapclientAt4min();
			msg = checkMailSize(msg);
			clientforimap_Output.appendMessage(iconnforimap_Output, msg);

		} 
		else if (EmailWizardApplication.r_gmail_app.isSelected()) {
			try {
				msg = checkMailSize(msg);
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
		
		
		else if (EmailWizardApplication.r_csv.isSelected()) {

			saveCSVEmailandTask(msg);
			msg.close();

		}

		else {
			note.save(destination.getAbsolutePath() + File.separator + mailCount + "_" + subjectName + ".msg",
					NoteSaveFormat.Msg);
			msg.close();
			note.close();
		}

	}

	private void downloadTask(Item itemType, int mailCount, Object value) throws Exception {

		if (value instanceof FolderInfo) {
			pstfolderInfo = (FolderInfo) value;
		} else if (value instanceof MboxrdStorageWriter) {
			mbox = (MboxrdStorageWriter) value;
		} else if (value instanceof File) {
			destination = (File) value;
		} else if (value instanceof Workbook) {
			workbook = (Workbook) value;
		}

		StreamingSubscriptionConnection conn = new StreamingSubscriptionConnection(service, 30);
		PropertySet propertySet = new PropertySet(TaskSchema.MimeContent);
		propertySet.setBasePropertySet(BasePropertySet.FirstClassProperties);
		Task task = Task.bind(service, itemType.getId(), propertySet);
		ByteArrayInputStream bis = new ByteArrayInputStream(task.getMimeContent().getContent());
		MailMessage msg = new MailMessage().load(bis);
		String subjectName = null;
		if (EmailWizardApplication.r_pst.isSelected()) {

			MapiTask mapiTask = MapiMessageUtils.convertToMapiTask(bis, task);
			subjectName = mapiTask.getSubject();
			mapiTask = MapiMessageUtils.setTaskMapiProperty(mapiTask, task);
			pstfolderInfo.changeContainerClass("IPF.Task");
			pstfolderInfo.addMapiMessageItem(mapiTask);
			mapiTask.close();
			bis.close();
		} else if (EmailWizardApplication.r_mbox.isSelected()) {

			subjectName = msg.getSubject();
			mbox.writeMessage(msg);
		} else if (EmailWizardApplication.r_gmail.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected() || EmailWizardApplication.r_yandex.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

			breakImapclientAt4min();
			clientforimap_Output.appendMessage(iconnforimap_Output, msg);

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
		
		
		else if (EmailWizardApplication.r_csv.isSelected()) {

			saveCSVEmailandTask(msg);
			msg.close();

		}

		else {

			subjectName = FileNamingUtils.namingConvention(msg.getSubject());
			msg.save(destination.getAbsolutePath() + File.separator + mailCount + "_" + subjectName + ".msg",
					SaveOptions.getDefaultMsg());
			msg.close();
		}

		EmailWizardApplication.downloadingFileName.setText(folderMailCount + "_" + subjectName);

	}

	public MailMessage checkMailSize(MailMessage msg) {

		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
		msg.save(emlStream, SaveOptions.getDefaultMsg());
		if (emlStream.size() > EmailWizardApplication.IMAP_MAIL_SIZE) {
			logger.warn("Mail over size for imap server: " + emlStream.size());
			msg.getAttachments().clear();
		}

		return msg;

	}

	private void downloadContact(Item itemType, int mailCount, Object value) throws ServiceLocalException, Exception {

		if (value instanceof FolderInfo) {
			pstfolderInfo = (FolderInfo) value;
		} else if (value instanceof MboxrdStorageWriter) {
			mbox = (MboxrdStorageWriter) value;
		} else if (value instanceof File) {
			destination = (File) value;
		} else if (value instanceof Workbook) {
			workbook = (Workbook) value;
		}

		PropertySet propertySet = new PropertySet(ContactSchema.MimeContent, ContactSchema.Attachments);
		propertySet.setBasePropertySet(BasePropertySet.FirstClassProperties);

		Contact contact = Contact.bind(service, itemType.getId(), propertySet);
		ByteArrayInputStream bis = new ByteArrayInputStream(contact.getMimeContent().getContent());
		MapiContact mapiContact = MapiContact.fromVCard(bis);
		mapiContact = MapiMessageUtils.setContactMapiProperty(mapiContact, contact);

		if (contact.getHasAttachments()) {
			microsoft.exchange.webservices.data.property.complex.AttachmentCollection attachments = contact
					.getAttachments();

			for (microsoft.exchange.webservices.data.property.complex.Attachment attachment : attachments) {
				FileAttachment fileAttachment = (FileAttachment) attachment;
				fileAttachment.load();
				System.out.println(attachment.getName());
				mapiContact.getAttachments().add(attachment.getName(), fileAttachment.getContent());

			}

		}

		String subjectName = FileNamingUtils.namingConvention(contact.getDisplayName());
		if (EmailWizardApplication.r_pst.isSelected()) {

			pstfolderInfo.changeContainerClass("IPF.Contact");
			pstfolderInfo.addMapiMessageItem(mapiContact);
			bis.close();
		} else if (EmailWizardApplication.r_mbox.isSelected()) {
			ByteArrayInputStream biss = new ByteArrayInputStream(contact.getMimeContent().getContent());
			MailMessage msg = new MailMessage().load(biss);
			mbox.writeMessage(msg);
			msg.close();
		}

		else if (EmailWizardApplication.r_gmail.isSelected() 
				|| EmailWizardApplication.r_aol.isSelected()
				||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()
				||EmailWizardApplication.r_hotmail.isSelected() || EmailWizardApplication.r_yandex.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

			breakImapclientAt4min();

			ByteArrayOutputStream os = new ByteArrayOutputStream();
			mapiContact.save(os, ContactSaveFormat.VCard);
			MailMessage mailMessage = new MailMessage();
			Attachment attachment = new Attachment(new ByteArrayInputStream(os.toByteArray()), subjectName + ".vcf");
			mailMessage.addAttachment(attachment);
			clientforimap_Output.appendMessage(iconnforimap_Output, mailMessage);
			os.close();
			mailMessage.close();
			mapiContact.close();

		}
		
		else if (EmailWizardApplication.r_gmail_app.isSelected()) {
			MailMessage mailMessage = new MailMessage();
			try {
				ByteArrayOutputStream os = new ByteArrayOutputStream();
				mapiContact.save(os, ContactSaveFormat.VCard);				
				Attachment attachment = new Attachment(new ByteArrayInputStream(os.toByteArray()), subjectName + ".vcf");
				mailMessage.addAttachment(attachment);
				gmailImport(mailMessage);
				os.close();
				mailMessage.close();
				mapiContact.close();
				
											
				} catch (Exception exception) {
					logger.error(exception.getMessage());
					ExceptionHandler exceptionHandler=new ExceptionHandler(exception);						
					if(exceptionHandler.GsuiteExceptionHandler())
					{
						EmailWizardApplication.lblNoInternetConnection.setVisible(true);
						while (!checkInternet()) {}
						EmailWizardApplication.lblNoInternetConnection.setVisible(false);
						getOutputGmailAppService();
						gmailImport(mailMessage);
	
					}
		}
		}

		else if (EmailWizardApplication.r_csv.isSelected()) {

			saveCSVContact(mapiContact);
		}

		else {

			Files.write(contact.getMimeContent().getContent(),
					new File(destination.getAbsolutePath() + File.separator + mailCount + "_" + subjectName + ".vcf"));

		}

		mapiContact.close();

		EmailWizardApplication.downloadingFileName.setText(folderMailCount + "_" + subjectName);
	}

	private void downloadAppointment(Item itemType, int mailCount, Object value)
			throws ServiceLocalException, Exception {
		if (value instanceof FolderInfo) {
			pstfolderInfo = (FolderInfo) value;
		} else if (value instanceof MboxrdStorageWriter) {
			mbox = (MboxrdStorageWriter) value;
		} else if (value instanceof File) {
			destination = (File) value;
		} else if (value instanceof Workbook) {
			workbook = (Workbook) value;
		}

		PropertySet propertySet = getAppoinmentPropertySet();
		Appointment appoinment = Appointment.bind(service, itemType.getId(), propertySet);
		MapiCalendar mapiCalendar = getMapiCalendar(appoinment);
		String subjectName = mapiCalendar.getSubject();

		if (EmailWizardApplication.r_pst.isSelected()) {

			pstfolderInfo.changeContainerClass("IPF.Appointment");
			pstfolderInfo.addMapiMessageItem(mapiCalendar);

		} else if (EmailWizardApplication.r_mbox.isSelected()) {
			ByteArrayInputStream bis = new ByteArrayInputStream(appoinment.getMimeContent().getContent());
			MailMessage msg = new MailMessage().load(bis);
			subjectName = msg.getSubject();
			if (appoinment.getHasAttachments()) {
				microsoft.exchange.webservices.data.property.complex.AttachmentCollection attachments = appoinment
						.getAttachments();

				for (microsoft.exchange.webservices.data.property.complex.Attachment attachment : attachments) {

					FileAttachment fileAttachment = (FileAttachment) attachment;
					fileAttachment.load();
					InputStream targetStream = new ByteArrayInputStream(fileAttachment.getContent());
					Attachment att = new Attachment(attachment.getName());
					att.setContentStream(targetStream);
					msg.getAttachments().addItem(att);
					targetStream.close();
				}

			}

			mbox.writeMessage(msg);
		} else if (EmailWizardApplication.r_gmail.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected() || EmailWizardApplication.r_yandex.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

			breakImapclientAt4min();
			ByteArrayInputStream bis = new ByteArrayInputStream(appoinment.getMimeContent().getContent());

			MailMessage msg = new MailMessage();
			msg.setSubject(subjectName);
			MailAddressCollection toMailAddress = new MailAddressCollection();
			toMailAddress.addMailAddress(new MailAddress(appoinment.getDisplayTo(), true));
			com.aspose.email.Appointment target = new com.aspose.email.Appointment(appoinment.getLocation(),
					appoinment.getStart(), appoinment.getEnd(), new MailAddress(appoinment.getInReplyTo(), true),
					toMailAddress);
			target.setSummary(appoinment.getSubject());
			target.setDescription(appoinment.getBody().toString());
			msg.addAlternateView(target.requestApointment());

			if (appoinment.getHasAttachments()) {
				microsoft.exchange.webservices.data.property.complex.AttachmentCollection attachments = appoinment
						.getAttachments();

				for (microsoft.exchange.webservices.data.property.complex.Attachment attachment : attachments) {

					FileAttachment fileAttachment = (FileAttachment) attachment;
					fileAttachment.load();
					InputStream targetStream = new ByteArrayInputStream(fileAttachment.getContent());
					Attachment att = new Attachment(targetStream, attachment.getContentType());
					att.setName(attachment.getName());
					msg.getAttachments().addItem(att);
					targetStream.close();
				}

			}

			clientforimap_Output.appendMessage(iconnforimap_Output, msg);
			bis.close();

		} 
		else if (EmailWizardApplication.r_gmail_app.isSelected()) {
			MailMessage msg = new MailMessage();
			try {
				ByteArrayInputStream bis = new ByteArrayInputStream(appoinment.getMimeContent().getContent());				
				msg.setSubject(subjectName);
				MailAddressCollection toMailAddress = new MailAddressCollection();
				toMailAddress.addMailAddress(new MailAddress(appoinment.getDisplayTo(), true));
				com.aspose.email.Appointment target = new com.aspose.email.Appointment(appoinment.getLocation(),
						appoinment.getStart(), appoinment.getEnd(), new MailAddress(appoinment.getInReplyTo(), true),
						toMailAddress);
				target.setSummary(appoinment.getSubject());
				target.setDescription(appoinment.getBody().toString());
				msg.addAlternateView(target.requestApointment());

				if (appoinment.getHasAttachments()) {
					microsoft.exchange.webservices.data.property.complex.AttachmentCollection attachments = appoinment
							.getAttachments();

					for (microsoft.exchange.webservices.data.property.complex.Attachment attachment : attachments) {

						FileAttachment fileAttachment = (FileAttachment) attachment;
						fileAttachment.load();
						InputStream targetStream = new ByteArrayInputStream(fileAttachment.getContent());
						Attachment att = new Attachment(targetStream, attachment.getContentType());
						att.setName(attachment.getName());
						msg.getAttachments().addItem(att);
						targetStream.close();
					}

				}
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
		
		else if (EmailWizardApplication.r_csv.isSelected()) {
			workbook = CSVUtils.saveCSVAppointment(workbook, cellNo, mapiCalendar);
			cellNo++;
		}

		else {

			subjectName = FileNamingUtils.namingConvention(appoinment.getSubject());
			mapiCalendar.save(destination.getAbsolutePath() + File.separator + mailCount + "_" + subjectName + ".ics",
					AppointmentSaveFormat.Ics);

		}

		EmailWizardApplication.downloadingFileName.setText(folderMailCount + "_" + subjectName);
	}

	public PropertySet getAppoinmentPropertySet() {
		PropertySet propertySet = new PropertySet(AppointmentSchema.MimeContent, AppointmentSchema.Attachments,
				AppointmentSchema.HasAttachments, AppointmentSchema.DisplayTo, AppointmentSchema.Location,
				AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Subject,
				AppointmentSchema.DateTimeCreated, AppointmentSchema.Body, AppointmentSchema.ItemClass,
				AppointmentSchema.InReplyTo, AppointmentSchema.LastModifiedTime, AppointmentSchema.DateTimeReceived);
		propertySet.setRequestedBodyType(BodyType.Text);
		return propertySet;
	}

	public MapiCalendar getMapiCalendar(Appointment appoinment) throws Exception {
		MapiCalendar mapiCalendar = new MapiCalendar(appoinment.getLocation(), appoinment.getSubject(),
				appoinment.getBody().toString(), appoinment.getStart(), appoinment.getEnd());
		mapiCalendar = MapiMessageUtils.setCalendarMapiProperty(mapiCalendar, appoinment);

		if (appoinment.getHasAttachments()) {
			microsoft.exchange.webservices.data.property.complex.AttachmentCollection attachments = appoinment
					.getAttachments();

			for (microsoft.exchange.webservices.data.property.complex.Attachment attachment : attachments) {

				FileAttachment fileAttachment = (FileAttachment) attachment;
				fileAttachment.load();
				mapiCalendar.getAttachments().add(attachment.getName(), fileAttachment.getContent());

			}

		}
		return mapiCalendar;
	}

	private void downloadEmails(MailMessage msg, String subjectName, int mailCount, Object value) throws Exception {

		if (value instanceof FolderInfo) {
			pstfolderInfo = (FolderInfo) value;
		} else if (value instanceof MboxrdStorageWriter) {
			mbox = (MboxrdStorageWriter) value;
		} else if (value instanceof File) {
			destination = (File) value;
		} else if (value instanceof String) {
			imapFolderPath = (String) value;
		} else if (value instanceof Workbook) {
			workbook = (Workbook) value;
		}

		if (EmailWizardApplication.r_pst.isSelected()) {
			pstfolderInfo.addMessage(MapiMessage.fromMailMessage(msg));
			msg.close();
		} else if (EmailWizardApplication.r_mbox.isSelected()) {

			mbox.writeMessage(msg);
		}

		else if (EmailWizardApplication.r_csv.isSelected()) {

			saveCSVEmailandTask(msg);
		} else if (EmailWizardApplication.r_gmail.isSelected() || EmailWizardApplication.r_aol.isSelected()||EmailWizardApplication.r_aws.isSelected()
				|| EmailWizardApplication.r_office.isSelected() || EmailWizardApplication.r_yahoo.isSelected()
				|| EmailWizardApplication.r_zoho.isSelected()|| EmailWizardApplication.r_icloud.isSelected()||EmailWizardApplication.r_hotmail.isSelected() || EmailWizardApplication.r_yandex.isSelected()
				|| EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

			breakImapclientAt4min();
			msg = checkMailSize(msg);
			clientforimap_Output.appendMessage(iconnforimap_Output, msg);

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

		else {
			String destinationPath = destination.getAbsolutePath() + File.separator + mailCount + "_" + subjectName;
			DownloadWordFormat dwf = new DownloadWordFormat();
			dwf.saveWordFormat(msg, destinationPath);
		}
	}

	public Document convertMSGToDocument(MailMessage msg) throws Exception {
		LoadOptions lo = new LoadOptions();
		lo.setLoadFormat(LoadFormat.MHTML);
		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
		msg.save(emlStream, SaveOptions.getDefaultMhtml());

		Document document = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
		emlStream.close();
		msg.close();
		return document;
	}

	private void breakImapclientAt4min() {
		if (EmailWizardApplication.checkImapConnectionTime <= System.currentTimeMillis()) {
			EmailWizardApplication.lblNoInternetConnection.setText("Imap Connection Refresh");
			LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Doing refresh for Imap Server");	
			imapOutputConnectionAgain();
			
			EmailWizardApplication.checkImapConnectionTime = System.currentTimeMillis() + EmailWizardApplication.IMAP_RERESH_TIMEOUT;
			LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Refresh done");

		}
	}

	private void downloadSubFolder(ExchangeService service, FolderId folderID, File destination) throws Exception {

		FindFoldersResults findFolderResults = service.findFolders(folderID, new FolderView(Integer.MAX_VALUE));
		for (Folder folder : findFolderResults) {

			if (folder.getChildFolderCount() > 0) {
				File finalDestination = createFolder(folder, destination);
				downloadFolder(folder, finalDestination);
				downloadSubFolder(service, folder.getId(), finalDestination);

			} else {
				File finalDestination = createFolder(folder, destination);
				downloadFolder(folder, finalDestination);

			}

		}
	}

	private void downloadSubFolderForCSV(ExchangeService service, FolderId folderID, File destination)
			throws Exception {

		FindFoldersResults findFolderResults = service.findFolders(folderID, new FolderView(Integer.MAX_VALUE));
		for (Folder folder : findFolderResults) {

			if (folder.getChildFolderCount() > 0) {

				File finalDestination = createFolder(folder, destination);
				Workbook workbook = null;
				try {
					if (folder.getFolderClass().contains("IPF.Contact")) {
						workbook = createCSVStructureContact();
					} else {
						workbook = createCSVStructure();
					}					
				} 
				catch (Exception e) {
					workbook = createCSVStructure();
					
				}finally {
					downloadFolder(folder, workbook);
					saveCSV(workbook, finalDestination);
				}

				downloadSubFolderForCSV(service, folder.getId(), finalDestination);
			} else {
				File finalDestination = createFolder(folder, destination);

				Workbook workbook = null;
				try {
					if (folder.getFolderClass().contains("IPF.Contact")) {
						workbook = createCSVStructureContact();
					} else {
						workbook = createCSVStructure();
					}					
				}
				catch (Exception e) {
					workbook = createCSVStructure();
					
				}finally {
					downloadFolder(folder, workbook);
					saveCSV(workbook, finalDestination);
				}

			}

		}
	}
	private void downloadSubFolderForImapClient(ExchangeService service, FolderId folderID, String destination)
			throws Exception {

		FindFoldersResults findFolderResults = service.findFolders(folderID, new FolderView(Integer.MAX_VALUE));
		for (Folder folder : findFolderResults) {

			if (folder.getChildFolderCount() > 0) {
                  
				String folderPathImap = createFolderInImap(folder, destination);
				downloadFolder(folder, folderPathImap);
				downloadSubFolderForImapClient(service, folder.getId(), folderPathImap);
			} else {
				 System.out.println(folder.getDisplayName());
				String folderPathImap = createFolderInImap(folder, destination);
				downloadFolder(folder, folderPathImap);

			}

		}
	}
	private void downloadSubFolderForGmailApp(ExchangeService service, FolderId folderID, String destination)throws Exception {
			
		FindFoldersResults findFolderResults = service.findFolders(folderID, new FolderView(Integer.MAX_VALUE));
		for (Folder folder : findFolderResults) {

			if (folder.getChildFolderCount() > 0) {
                  
				String folderPathImap = createFolderGmailApp(folder, destination);
				downloadFolder(folder, folderPathImap);
				downloadSubFolderForGmailApp(service, folder.getId(), folderPathImap);
			} else {
				String folderPathImap  = createFolderGmailApp(folder, destination);
				 downloadFolder(folder, folderPathImap);

			}

		}
	}

	private void downloadSubFolderForMBOX(ExchangeService service, FolderId folderID, File destination)
			throws Exception {

		FindFoldersResults findFolderResults = service.findFolders(folderID, new FolderView(Integer.MAX_VALUE));
		for (Folder folder : findFolderResults) {

			if (folder.getChildFolderCount() > 0) {

				File finalDestination = createMBOX(folder, destination);
				mbox = new MboxrdStorageWriter(
						finalDestination.getAbsolutePath() + File.separator + finalDestination.getName() + ".mbx",
						false);
				downloadFolder(folder, mbox);
				downloadSubFolderForMBOX(service, folder.getId(), finalDestination);
			} else {
				File finalDestination = createMBOX(folder, destination);
				mbox = new MboxrdStorageWriter(
						finalDestination.getAbsolutePath() + File.separator + finalDestination.getName() + ".mbx",
						false);
				downloadFolder(folder, mbox);

			}

		}
	}

	private void downloadSubFolderForPST(ExchangeService service, Folder folder, String pstPath) throws Exception {

		FindFoldersResults findFolderResults = service.findFolders(folder.getId(), new FolderView(Integer.MAX_VALUE));
		for (Folder searchfolder : findFolderResults) {
			if (searchfolder.getChildFolderCount() > 0) {

				String childFolderPath = pstPath + File.separator + searchfolder.getDisplayName();

				FolderInfo folderInfo = pst.getRootFolder().addSubFolder(childFolderPath, true);
				downloadFolder(searchfolder, folderInfo);
				downloadSubFolderForPST(service, searchfolder, childFolderPath);
				childFolderPath = pstPath;

			} else {

				String noChildFolder = pstPath + File.separator + searchfolder.getDisplayName();
				FolderInfo noChildfolderInfo = pst.getRootFolder().addSubFolder(noChildFolder, true);
				downloadFolder(searchfolder, noChildfolderInfo);

			}

		}
	}

	private String createFolderInImap(Folder folder, String destination) throws ServiceLocalException {
		if (EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {
			return createFolderImapClient(folder, destination);
		} else {
			return createFolderEmailClient(folder, destination);
		}

	}
	
	private String createFolderGmailApp(Folder folder, String destination) throws IOException, ServiceLocalException {

		String folderPathImap = destination + "/" + FileNamingUtils.buildImapFolderName(folder.getDisplayName());
		Label label = new Label();
		label.setName(folderPathImap);
		label.setLabelListVisibility("labelShow");
		label.setMessageListVisibility("show");
		parentLabel = outputGmailService.users().labels().create(userName, label).execute();	
		return folderPathImap;
	}

	private String createFolderEmailClient(Folder folder, String destination) throws ServiceLocalException {

		String folderPathImap = destination + "/" + FileNamingUtils.buildImapFolderName(folder.getDisplayName());
		try
		{
			clientforimap_Output.createFolder(iconnforimap_Output,folderPathImap);
		}
		catch (ImageException e) {
			// TODO: handle exception
			if(e.getMessage().contains("AE_8_2_0058 NO [CANNOT] Folder name is not allowed."))
			{
				LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"The Folder name is too long. Please try another name:");
			}
		}

		return folderPathImap;
	}

	private String createFolderImapClient(Folder folder, String destination) throws ServiceLocalException {

		String folderName = FileNamingUtils.buildImapFolderName(folder.getDisplayName()).replace(".", "-");

		String[] split = folderName.split("/");
		String parentPath = destination;
		String removeDotlabelName =null;
		String removeSalshWithDotlabelName=null;
		for (String string : split) {
			 removeDotlabelName =string.replace(".", "-");
			 removeSalshWithDotlabelName =removeDotlabelName.replace("/", ".");
			 string = FileNamingUtils.getRidOfIllegalFileNameCharacters(removeSalshWithDotlabelName);
			String folderPathImap = parentPath + "." + string;
			if (!clientforimap_Output.existFolder("INBOX." + folderPathImap)) {
				clientforimap_Output.createFolder(iconnforimap_Output, "INBOX." + folderPathImap);
			}
			 removeDotlabelName =string.replace(".", "-");
			 removeSalshWithDotlabelName =removeDotlabelName.replace("/", ".");
			 string = FileNamingUtils.getRidOfIllegalFileNameCharacters(removeSalshWithDotlabelName);
			parentPath = parentPath + "." + string;
			clientforimap_Output.selectFolder(iconnforimap_Output, "INBOX." + folderPathImap);
			clientforimap_Output.subscribeFolder(iconnforimap_Output, "INBOX." + folderPathImap);
		}

		return parentPath;
	}

	private File createMBOX(Folder folder, File destination) throws ServiceLocalException {
		File destinationPath = new File(destination.getAbsolutePath() + File.separator
				+ FileNamingUtils.getRidOfIllegalFileNameCharacters(folder.getDisplayName().trim()));
		destinationPath.mkdirs();
		return destinationPath;
	}

	private File createFolder(Folder folder, File destination) throws ServiceLocalException {
		File destinationPath = new File(destination.getAbsolutePath() + File.separator
				+ FileNamingUtils.getRidOfIllegalFileNameCharacters(folder.getDisplayName().trim()));
		destinationPath.mkdirs();
		return destinationPath;
	}

	public ExchangeService connectionReset() {
		return service = EmailWizardApplication.resetEWS();
	}

	public String isNamingConventionSelected(MailMessage msgFull) {
		String subjectName = null;
		if (EmailWizardApplication.chckbxNamingconvention.isSelected()) {
			subjectName = FileNamingUtils.buildFileName(msgFull, msgFull.getDate().getTime());
		} else {
			try {
				subjectName = FileNamingUtils.namingConvention(msgFull.getSubject());
			} catch (NoSuchElementException e) {

				logger.error("An exception occurred!", e);
			}

		}

		return subjectName;
	}

	public FindItemsResults<Item> getSearchResults(Folder folder, ItemView itemView) throws Exception {
		if (EmailWizardApplication.rdbtnDateFilter.isSelected()) {
			SearchFilter dateFilter = queryDateExist(folder.getId());
			return service.findItems(folder.getId(), dateFilter, itemView);
		}
		return service.findItems(folder.getId(), itemView);

	}

	public SearchFilter queryDateExist(FolderId folderID) {

		Calendar calendarstartdate = EmailWizardApplication.start_dateChooser.getCalendar();
		calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
		calendarstartdate.set(Calendar.MINUTE, 00);
		calendarstartdate.set(Calendar.SECOND, 00);

		Calendar calendarenddate = EmailWizardApplication.end_dateChooser.getCalendar();
		calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
		calendarenddate.set(Calendar.MINUTE, 59);
		calendarenddate.set(Calendar.SECOND, 59);

		SearchFilter greaterThanfilter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived,
				calendarstartdate.getTime());
		SearchFilter lessThanfilter = new SearchFilter.IsLessThan(ItemSchema.DateTimeReceived,
				calendarenddate.getTime());
		SearchFilter dateFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, greaterThanfilter,
				lessThanfilter);

		return dateFilter;

	}

	public boolean ischeckDateExist(Long emailMilliseconds) {

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

	public void imapOutputConnectionAgain() {

		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Please wait....Trying To Reconnect");
		EmailWizardApplication.lblNoInternetConnection.setVisible(true);
		if (EmailWizardApplication.clientforimap_Output != null && EmailWizardApplication.iconnforimap_Output != null) {

			EmailWizardApplication.clientforimap_Output.close();
			EmailWizardApplication.clientforimap_Output.dispose();

			EmailWizardApplication.iconnforimap_Output.dispose();
			if (!EmailWizardApplication.iconnforimap_Output.isDisposed()) {
				EmailWizardApplication.clientforimap_Output.dispose();
			}
		}
		while (true) {
			try {
				clientforimap_Output = EmailWizardApplication.outputImapConnection();
				iconnforimap_Output = clientforimap_Output.createConnection();
				if (EmailWizardApplication.r_imap.isSelected()||EmailWizardApplication.r_hostgator.isSelected()) {

					clientforimap_Output.selectFolder(iconnforimap_Output, "INBOX." + imapFolderPath);
				} 
//				else if(EmailWizardApplication.r_aws.isSelected())
//				{
//					clientforimap_Output.selectFolder(iconnforimap_Output, imapFolderPath);	
//				}
//				else
//				{
//					clientforimap_Output.subscribeFolder(iconnforimap_Output, imapFolderPath);
//				}
				

				break;
			} catch (Exception e) {

			}
		}
		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Connected to Imap server");
		EmailWizardApplication.lblNoInternetConnection.setVisible(false);

	}

	public void setUpProgressBar(Folder folder) {
		try {
			folderMailCount = 0;
			duplicateEmailsList = new ArrayList<String>();
			EmailWizardApplication.modelDownloading.setValueAt(folder.getTotalCount(), EmailWizardApplication.rownCount, 4);
			EmailWizardApplication.progressBar_Downloading.setValue(0);
			EmailWizardApplication.progressBar_Downloading.setVisible(true);
			EmailWizardApplication.progressBar_Downloading.setMaximum(100);
			EmailWizardApplication.progressBar_Downloading.setStringPainted(true);
			EmailWizardApplication.lblDownloading.setVisible(true);
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public boolean checkInternetconnection(ServiceRequestException ex, FolderId folderID, ItemView itemView) {

		ex.printStackTrace();
		boolean check = true;
		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Please wait....Trying To Reconnect");
		EmailWizardApplication.lblNoInternetConnection.setVisible(true);
		while (check) {
			if (ex.getMessage().contains("The request failed. outlook.office365.com")
					|| ex.getMessage().contains("The request failed. The request failed. Connection reset")
					|| ex.getMessage().contains("Connection reset") || ex.getMessage().contains("outlook.office365.com")
					|| ex.getMessage().contains("The request failed. java.net.SocketException: Connection reset")) {
				try {
					service.findItems(folderID, itemView);
					check = false;
				} catch (ServiceRequestException serviceRequestException) {
					ex = serviceRequestException;
					check = true;
				} catch (Exception e) {

					logger.error("An exception occurred!", e);

				}

			} else if (ex.getMessage().contains("The request failed. The remote server returned an error: (401)")) {
				service = EmailWizardApplication.loginWithRefreshTokenEWS(EmailWizardApplication.input_userName);
				check = false;
			} else {
				logger.error("An exception occurred!", ex);
				check = true;
				break;
			}

		}
		LogUtils.setTextToLogScreen(EmailWizardApplication.textPane_log,logger,"Connected to server");
		EmailWizardApplication.lblNoInternetConnection.setVisible(false);
		return check;
	}

	public Workbook createCSVStructure() {
		cellNo = 1;
		Workbook workbook = CSVUtils.createCSVStructure(cellNo);
		cellNo++;
		return workbook;
	}

	public Workbook createCSVStructureContact() {
		cellNo = 1;
		Workbook workbook = CSVUtils.createCSVStructureContact(cellNo);
		cellNo++;
		return workbook;
	}
	
	public void saveCSVEmailandTask(MailMessage msg) {
		workbook = CSVUtils.saveCSVEmailandTask(workbook, cellNo, msg);
		cellNo++;
	}
	public void saveCSVContact(MapiContact mapiContact) {
		workbook = CSVUtils.saveCSVContact(workbook, cellNo, mapiContact);
		cellNo++;
	}
	private void saveCSV(Workbook workbook, File finalDestination) {
		workbook=CSVUtils.saveCSV(workbook,finalDestination);
		workbook.dispose();
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
			String pstFileName=pst.getStore().getDisplayName();
			pst.close();
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH-mm-ss");
			LocalDateTime now = LocalDateTime.now();			
			EmailWizardApplication.pst = PersonalStorage.create(EmailWizardApplication.pstSplitFile.getParent() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + pstFileName + ".pst", FileFormatVersion.Unicode);					
			EmailWizardApplication.pstSplitFile = new File(EmailWizardApplication.pstSplitFile.getParent() + File.separator + "(" + splitCount + ") "+ dtf.format(now) + "-" + pstFileName + ".pst");					
			EmailWizardApplication.pst.getStore().changeDisplayName(pstFileName);			
			String labelBackwordSlash = validateImapFolderName(folderPath).replace("/", "\\");
			setPst();
			EmailWizardApplication.pstfolderInfo = EmailWizardApplication.pst.getRootFolder().addSubFolder(labelBackwordSlash, true);
			System.out.println(pstfolderInfo.getDisplayName());
			EmailWizardApplication.splitCount++;
		}

	}
	public String validateImapFolderName(String folderPath)
	{		
		if (EmailWizardApplication.selectedInput.equals(InputSource.IMAP.getValue())||EmailWizardApplication.selectedInput.equals(InputSource.HOSTGATOR.getValue())) {
			folderPath=folderPath.replace("INBOX.", "").replace(".", "/");
			
		}
		return FileNamingUtils.validFileNameForWindows(folderPath);
	}
	public void setPst()
	{
		this.pst=EmailWizardApplication.pst;
		this.splitCount=EmailWizardApplication.splitCount;
		
	}
	public Gmail getOutputGmailAppService() throws GeneralSecurityException, IOException
	{
	  	  HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
	  	 GoogleLogin googleLogin = new GoogleLogin();
	  	 outputGmailService = new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(googleLogin.getoathGoogleCredential())).setApplicationName(APPLICATION_NAME).build();
		 return outputGmailService;		
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
}
