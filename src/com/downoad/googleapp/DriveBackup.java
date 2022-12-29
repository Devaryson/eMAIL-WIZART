package com.downoad.googleapp;

import java.awt.HeadlessException;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

import javax.swing.table.DefaultTableModel;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.FileList;
import com.main.EmailWizardApplication;

public class DriveBackup {

	static int i = 1;
	static int duplicateCount;
	static DefaultTableModel model;
	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	static GoogleCredential credentials;
	public static Drive driveService = null;

	static String serviceAccountUsers;
	static HttpRequestInitializer requestInitializer=null;
	
	@SuppressWarnings("deprecation")
	public void googleCredentials(String serviceAccountId, String serviceAccountUser, String p12File)throws GeneralSecurityException, IOException {
					
		serviceAccountUsers=serviceAccountUser;
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

		credentials = new GoogleCredential.Builder().setTransport(HTTP_TRANSPORT).setJsonFactory(JSON_FACTORY)
				.setServiceAccountId(serviceAccountId)
				.setServiceAccountScopes(Collections.singleton("https://www.googleapis.com/auth/drive"))
				.setServiceAccountUser(serviceAccountUser).setServiceAccountPrivateKeyFromP12File(new File(p12File))
				.build();
	
		    DriveBackup.requestInitializer=credentials;
		    driveService = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(requestInitializer)).setApplicationName(APPLICATION_NAME).build();

		    download(driveService, new File(EmailWizardApplication.detinationPath));
	}

	public void googleCredentials(Credential credentials) throws GeneralSecurityException, IOException
	{
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		
		 DriveBackup.requestInitializer=credentials;
		driveService = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(requestInitializer)).setApplicationName(APPLICATION_NAME).build();
				
		download(driveService, new File(EmailWizardApplication.detinationPath));
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

	public static FileList readFileAndFolder(FileList result, File folderName) throws GeneralSecurityException{

		EmailWizardApplication.progressBar_Downloading.setValue(0);
		EmailWizardApplication.progressBar_Downloading.setVisible(true);
		EmailWizardApplication.progressBar_Downloading.setMaximum(100);
		EmailWizardApplication.progressBar_Downloading.setStringPainted(true);
		EmailWizardApplication.downloadingFileName.setVisible(true);	

		List<com.google.api.services.drive.model.File> files = result.getFiles();
		int i = 1;
	
		for (int j = 0; j < files.size(); j++) {
			
			checkStopAndDemo(i);

			try {
				com.google.api.services.drive.model.File file = files.get(j);
				file = driveService.files().get(file.getId()).execute();
				System.out.println(file.getName());
				if (file.getMimeType().contentEquals("application/vnd.google-apps.folder")) {

					File subfolderName = new File(folderName.getAbsolutePath() + File.separator + validFileName(file.getName()));							
					subfolderName.mkdirs();
					result = driveService.files().list().setQ("'" + file.getId() + "'" + " in parents")
							.setPageSize(1000).execute();

					readFileAndFolder(result, subfolderName);

				} else {
					//file = driveService.files().get(file.getId()).execute();
					driveDatat(driveService, file, folderName);

				}
				int prog = (i * 100) / files.size();
				EmailWizardApplication.progressBar_Downloading.setValue(prog);
				i++;

			} catch (Exception e) {
				
				System.out.println(e.getMessage());
				if(e.getMessage().contains("416 Requested range not satisfiable"))
                 {
					System.out.println(e.getMessage());
                }
				else if (e.getMessage().contains("www.googleapis.com") 
						|| e.getMessage().contains("oauth2.googleapis.com")
						|| e.getMessage().contains("No route to host: connect")
						||e.getMessage().contains("Failed to refresh access token: Connection reset")
 						||e.getMessage().contains("Connection reset")
 						||e.getMessage().contains("Software caused connection abort: connect")
 						||e.getMessage().contains("Read timed out")) {
					     e.printStackTrace();
					
					EmailWizardApplication.lblNoInternetConnection.setVisible(true);
 					while (!checkInternet()) { 						
 					}
					EmailWizardApplication.lblNoInternetConnection.setVisible(false);
			
					try {
						NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
						driveService = new Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(requestInitializer)).setApplicationName(APPLICATION_NAME).build();
								
						j--;
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					
					}

				}else
				{
					e.printStackTrace();
			
				}
			} 
		}

		return null;
	}

	public static void download(Drive driveService, File folderName) throws GeneralSecurityException, IOException {

		i = 1;
		FileList result;
	
			result = driveService.files().list().setQ("'root' in parents").execute();
			readFileAndFolder(result, folderName);
		
//		GoogleMaineFrame.lblDownloading.setVisible(false);
//		GoogleMaineFrame.progressBar_Downloading.setVisible(false);
	}

	public static void driveDatat(Drive service, com.google.api.services.drive.model.File file, File folderName) throws IOException {

		BufferedOutputStream outputStream = null;
		FileOutputStream outStream = null;
		try {
			String fileId = file.getId();
			if (file.getMimeType().contains("application/pdf")) {

				String fileName = file.getName();
				if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)+ ".pdf").exists()) {
					fileName = duplicateCount + "_" + validFileName(fileName);
				}

				outStream = new FileOutputStream(
						folderName.getAbsolutePath() + File.separator +i+"_"+ validFileName(fileName) + ".pdf");
				outputStream = new BufferedOutputStream(outStream);

				driveService.files().get(fileId).executeMediaAndDownloadTo(outputStream);

			}

			else if (file.getMimeType().contains("application/vnd.google-apps.spreadsheet")) {

				String fileName = file.getName();
				if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)+ ".xlsx").exists()) {
					fileName = duplicateCount + "_" + validFileName(fileName);
					duplicateCount++;
				}

				outStream = new FileOutputStream(
						folderName.getAbsolutePath() + File.separator + validFileName(fileName) + ".xlsx");

				outputStream = new BufferedOutputStream(outStream);

				driveService.files().export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
						.executeMediaAndDownloadTo(outputStream);

			} else if (file.getMimeType().contains("application/vnd.google-apps.document")) {
				try {
					
					String fileName = file.getName();

					if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".docx").exists()) {
						fileName = duplicateCount + "_" + validFileName(fileName);
						duplicateCount++;
					}
					outStream = new FileOutputStream(
							folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".docx");
					outputStream = new BufferedOutputStream(outStream);
					driveService.files().export(fileId, "application/vnd.openxmlformats-officedocument.wordprocessingml.document").executeMediaAndDownloadTo(outputStream);

				} catch (Exception e) {
					// TODO: handle exception
					e.printStackTrace();
					StringWriter errors = new StringWriter();
					e.printStackTrace(new PrintWriter(errors));

					//Main_Frame.logger.warning(errors + System.lineSeparator());
				}
			}
			 else if (file.getMimeType().contains("application/vnd.google-apps.drawing")) {
					try {
						//"application/vnd.google-apps.document"
						String fileName = file.getName();
						if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".png").exists()) {
							fileName = duplicateCount + "_" + validFileName(fileName);
							duplicateCount++;
						}
						outStream = new FileOutputStream(
								folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".png");
						outputStream = new BufferedOutputStream(outStream);
						driveService.files().export(fileId, "image/png").executeMediaAndDownloadTo(outputStream);

					} catch (Exception e) {
						// TODO: handle exception
						e.printStackTrace();
						StringWriter errors = new StringWriter();
						e.printStackTrace(new PrintWriter(errors));

						//Main_Frame.logger.warning(errors + System.lineSeparator());
					}
				}
			 else if (file.getMimeType().contains("application/vnd.google-apps.form")) {
					try {
						//"application/vnd.google-apps.document"
						String fileName = file.getName();
						if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".pdf").exists()) {
							fileName = duplicateCount + "_" + validFileName(fileName);
							duplicateCount++;
						}
						outStream = new FileOutputStream(
								folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".pdf");
						outputStream = new BufferedOutputStream(outStream);
						driveService.files().export(fileId, "application/vnd.google-apps.freebird").executeMediaAndDownloadTo(outputStream);
						

					} catch (Exception e) {
						// TODO: handle exception
						e.printStackTrace();
						StringWriter errors = new StringWriter();
						e.printStackTrace(new PrintWriter(errors));

						//Main_Frame.logger.warning(errors + System.lineSeparator());
					}
				}
			 else if (file.getMimeType().contains("application/vnd.google-apps.presentation")) {
					try {
				
						String fileName = file.getName();
						if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".pptx").exists()) {
							fileName = duplicateCount + "_" + validFileName(fileName);
							duplicateCount++;
						}
						outStream = new FileOutputStream(
								folderName.getAbsolutePath() + File.separator + validFileName(fileName)+".pptx");
						outputStream = new BufferedOutputStream(outStream);
						driveService.files().export(fileId, "application/vnd.openxmlformats-officedocument.presentationml.presentation").executeMediaAndDownloadTo(outputStream);
						

					} catch (Exception e) {
						// TODO: handle exception
						e.printStackTrace();
						StringWriter errors = new StringWriter();
						e.printStackTrace(new PrintWriter(errors));

						//Main_Frame.logger.warning(errors + System.lineSeparator());
					}
				}
			
			
			else {
				String fileName = file.getName();
				if (new File(folderName.getAbsolutePath() + File.separator + validFileName(fileName)).exists()) {
					fileName = duplicateCount + "_" + validFileName(fileName);
					duplicateCount++;
				}
				outStream = new FileOutputStream(folderName.getAbsolutePath() + File.separator + validFileName(fileName));						
				outputStream = new BufferedOutputStream(outStream);
				driveService.files().get(fileId).executeMediaAndDownloadTo(outputStream);
			
			}
			
			EmailWizardApplication.downloadingFileName.setText(validFileName(file.getName()));			
			EmailWizardApplication.modelDownloading.setValueAt(i, EmailWizardApplication.rownCount, 1);	
			i++;
		} finally {

			if (outStream != null) {
				try {
					outputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				try {
					outStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			}

			System.gc();
		}

	}

	public void printFile(Drive service) {

		try {

			try {

				int i = 1;

				FileList result = driveService.files().list().setQ("'root' in parents").execute();

				List<com.google.api.services.drive.model.File> files = result.getFiles();

				for (com.google.api.services.drive.model.File file : files) {

					System.out.println("Title: " + file.getName());
					System.out.println("Id: " + file.getId());

					String fileType = null;
					if (file.getMimeType().contentEquals("application/vnd.google-apps.folder")) {

						fileType = "Folder";

					} else {
						fileType = "File";
					}

					i++;

				}

			} catch (HeadlessException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				StringWriter errors = new StringWriter();
				e.printStackTrace(new PrintWriter(errors));

			}

		} catch (IOException e) {
			System.out.println("An error occurred: " + e);
			StringWriter errors = new StringWriter();
			e.printStackTrace(new PrintWriter(errors));

		}
	}

	public static String validFileName(String fileName) {
		
		return fileName.replaceAll("[^a-zA-Z0-9\\.\\-]", "_").trim();

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
	
	public static boolean checkInternet()
	{
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

}