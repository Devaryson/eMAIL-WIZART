package com.downoad.googleapp;

import java.awt.CardLayout;
import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiElectronicAddress;
import com.aspose.email.MapiRecipientCollection;
import com.aspose.email.MapiRecipientType;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.calendar.Calendar;
import com.google.api.services.calendar.model.Event;
import com.google.api.services.calendar.model.Event.Organizer;
import com.google.api.services.calendar.model.EventAttachment;
import com.google.api.services.calendar.model.EventAttendee;
import com.google.api.services.calendar.model.Events;
import com.main.EmailWizardApplication;


public class CalenderBackup {

	static CardLayout card;
	
	public static Calendar calenderService;

	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	static GoogleCredential gSuitecredentials;
	public static Credential auth;
	@SuppressWarnings("deprecation")
	
	public void  downloadGsuiteCalendar(String serviceAccountId, String serviceAccountUser, String p12File)	throws GeneralSecurityException, IOException {
		
		getGsuiteCalenderService(serviceAccountId,  serviceAccountUser,  p12File, new File(EmailWizardApplication.detinationPath));
		downloadCalender(new File(EmailWizardApplication.detinationPath));
		   
	}
	public void downloadGmailAPPCalendar(Credential gMailAppCredentials)throws GeneralSecurityException, IOException {
		
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		calenderService = new Calendar.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(gMailAppCredentials)).setApplicationName(APPLICATION_NAME).build();							
		downloadCalender(new File(EmailWizardApplication.detinationPath));
		    
	}
	
	@SuppressWarnings("deprecation")
	public void getGsuiteCalenderService(String serviceAccountId, String serviceAccountUser, String p12File,File folderName)
	{
		try {
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			gSuitecredentials = new GoogleCredential.Builder().setTransport(HTTP_TRANSPORT).setJsonFactory(JSON_FACTORY)
					.setServiceAccountId(serviceAccountId)
					.setServiceAccountScopes(Collections.singleton("https://www.googleapis.com/auth/calendar"))
					.setServiceAccountUser(serviceAccountUser).setServiceAccountPrivateKeyFromP12File(new File(p12File))
					.build();
			calenderService = new Calendar.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(gSuitecredentials)).setApplicationName(APPLICATION_NAME)
					.build();

			
		} catch (GeneralSecurityException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	@SuppressWarnings("deprecation")
	public static void downloadCalender(File folderName ) {

		// List the next 10 events from the primary calendar.
		try {
			
			Events events = calenderService.events().list("primary").execute();
			List<Event> items = events.getItems();
			if (items.isEmpty()) {
				System.out.println("No upcoming events found.");
			} else {
				System.out.println("Upcoming events");
				int i = 1;

				EmailWizardApplication.progressBar_Downloading.setValue(0);
				EmailWizardApplication.progressBar_Downloading.setVisible(true);
				EmailWizardApplication.progressBar_Downloading.setMaximum(100);
				EmailWizardApplication.lblDownloading.setVisible(true);


				String summary = null;

				for (int j = 0; j < items.size(); j++) {

					try {

						 checkStopAndDemo(j);

						Event event = items.get(j);

						com.google.api.client.util.DateTime start =null;
						com.google.api.client.util.DateTime end =null;
						
						
						try
						{
						    start = event.getStart().getDateTime();
						}
						catch (Exception e) {
							// TODO: handle exception
							System.out.println(e.getMessage());
						}
						try
						{
							end =  event.getEnd().getDateTime();
						}
						catch (Exception e) {
							// TODO: handle exception
							System.out.println(e.getMessage());
						}

						
						if (start == null) {
							start = event.getStart().getDate();

						} 
						if (end == null) {
							end = event.getStart().getDate();
						}
				
						Date startDate=null;
						Date endDate=null;
						if (start == null) {
							Date dateObj = new Date();
							 startDate = new Date(dateObj.getTime());
							

						} 
						if (end == null) {
							
							Date dateObj = new Date();
							endDate = new Date(dateObj.getTime());
	
							 
						}
						
						List<EventAttendee> attendee = event.getAttendees();
						MapiRecipientCollection attendeeCollection = new MapiRecipientCollection();
						if (attendee != null) {
							for (EventAttendee att : attendee) {
								attendeeCollection.add(att.getEmail(), att.getDisplayName(), MapiRecipientType.MAPI_TO);
							}
						}
						Organizer orgz = event.getOrganizer();
						 startDate = new Date(start.getValue());
						 endDate = new Date(end.getValue());

						MapiElectronicAddress mail = null;
						if (orgz != null) {
							mail = new MapiElectronicAddress(orgz.getEmail());

						}

						MapiCalendar meeting = new MapiCalendar(event.getLocation(), event.getSummary(),
								event.getDescription(), startDate, endDate, mail, attendeeCollection);

						List<EventAttachment> attachment = event.getAttachments();

						if (attachment != null) {
							for (EventAttachment attachmnetFile : attachment) {

								System.out.println(attachmnetFile.getTitle());

							}
						}

						if (event.getSummary() != null) {

							summary = getRidOfIllegalFileNameCharacters(event.getSummary());
							meeting.save(folderName.getAbsolutePath() + File.separator + i + summary + ".ics",
									AppointmentSaveFormat.Ics);
						} else {
							summary = " ";
							meeting.save(folderName.getAbsolutePath() + File.separator + i + ".ics",
									AppointmentSaveFormat.Ics);
						}

						meeting.close();

						System.out.println("downloaded");

						EmailWizardApplication.downloadingFileName.setText(summary+"_"+i);
						EmailWizardApplication.modelDownloading.setValueAt(i, EmailWizardApplication.rownCount, 3);

						int prog = (i * 100) / items.size();
						EmailWizardApplication.progressBar_Downloading.setValue(prog);

						i++;

					}

					catch (Exception e) {
						// TODO: handle exception
						e.printStackTrace();
						StringWriter errors = new StringWriter();
						e.printStackTrace(new PrintWriter(errors));

						EmailWizardApplication.logger.warn(errors + System.lineSeparator());
					}
				}
				EmailWizardApplication.lblDownloading.setVisible(false);
				EmailWizardApplication.progressBar_Downloading.setVisible(false);
			}

		} catch (Exception e) {

			e.printStackTrace();
			if(e.getMessage().contains("The user must be signed up for Google Calendar"))
			{
				return;
			}
			if (e.getMessage().contains("www.googleapis.com")
					|| e.getMessage().contains("oauth2.googleapis.com")
					|| e.getMessage().contains("No route to host: connect")
					||e.getMessage().contains("Failed to refresh access token: Connection reset")
					||e.getMessage().contains("Connection reset")
					||e.getMessage().contains("Software caused connection abort: connect")) {

				EmailWizardApplication.lblNoInternetConnection.setVisible(true);
				System.out.println("No Internet Connection");
					while (!checkInternet()) { 						
					}
				EmailWizardApplication.lblNoInternetConnection.setVisible(false);
				
             downloadCalender(folderName ) ;

			}
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
			System.out.println("Internet is not connected");
		} catch (IOException e) {
			System.out.println("Internet is not connected");
		}
		return false;
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
	private static HttpRequestInitializer setHttpTimeout(final HttpRequestInitializer requestInitializer) {
        return new HttpRequestInitializer() {
            @Override
            public void initialize(HttpRequest httpRequest) throws IOException {
                requestInitializer.initialize(httpRequest);
                httpRequest.setConnectTimeout(3 * 60000);  // 3 minutes connect timeout
                httpRequest.setReadTimeout(3 * 60000);  // 3 minutes read timeout
            }

	
        };
    }
	static boolean checkStopAndDemo(int i)
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
	
	

}
