package com.api.google;

import java.awt.CardLayout;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigInteger;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import javax.swing.ImageIcon;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;

import org.slf4j.LoggerFactory;

import com.google.api.client.auth.oauth2.StoredCredential;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.DataStore;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.admin.directory.Directory;
import com.google.api.services.admin.directory.DirectoryScopes;
import com.google.api.services.admin.directory.model.User;
import com.google.api.services.admin.directory.model.Users;
import com.google.auth.Credentials;
import com.google.auth.oauth2.UserCredentials;
import com.main.EmailWizardApplication;
import com.tool.info.ToolDetails;
import com.google.api.client.auth.oauth2.Credential;


@SuppressWarnings("deprecation")

public class GoogleLogin {
	
	public static org.slf4j.Logger logger=LoggerFactory.getLogger(EmailWizardApplication.class);
    final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String APPLICATION_NAME = "Gmail Backup";
	private static final String OATH_APPLICATION_NAME = "Aryson gmail backup tool";
    private GoogleCredential credential;
    private  Credentials oathCredentials;
    private  Credential oathCredential;
    private	String clientId = null;
    private	String clientSecret = null;
  
	public void googleCredentials(String serviceAccountId,String serviceAccountUser,String p12File)
	{
		try
		{
			 final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();	
			 GoogleCredential credential = new GoogleCredential.Builder()
                 .setTransport(HTTP_TRANSPORT)                                                    
                 .setJsonFactory(JSON_FACTORY)
                 .setServiceAccountId(serviceAccountId)
                 .setServiceAccountScopes(Collections.singleton(DirectoryScopes.ADMIN_DIRECTORY_USER_READONLY))
                 .setServiceAccountUser(serviceAccountUser)
                 .setServiceAccountPrivateKeyFromP12File(new File(p12File))
                 .build();
			      setGoogleCredential(credential);
		          getUserDetails(credential);
		          
		}
		catch (GeneralSecurityException | IOException e) {
	
			JOptionPane.showMessageDialog(null, 
					"User authorization failed (Access_denied) Please check you entered details!!",
					ToolDetails.messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
			
			logger.error("Error", e);
		}
	}

	public void getUserDetails(GoogleCredential credential) throws GeneralSecurityException, IOException {

		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		DefaultTableModel model = (DefaultTableModel) EmailWizardApplication.table_UserDetails.getModel();

		Directory servicedir = new Directory.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME).build();
		com.google.api.services.admin.directory.Directory.Users.List ul = servicedir.users().list()
				.setCustomer("my_customer");

		int userCount = 0;
		do {
			Users result = ul.execute();
			if (result != null && result.getUsers().size() > 0) {
				for (User user : result.getUsers()) {
					model.addRow(new Object[] { userCount, user.getName().getFullName(), user.getPrimaryEmail(), true });							
					userCount++;
				}
			} else {
				
				logger.info("No users found.");
				break;
			}
			ul.setPageToken(result.getNextPageToken());
		} while (ul.getPageToken() != null && ul.getPageToken().length() > 0);

		CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
		card.show(EmailWizardApplication.CardLayout, "GoogleUserDetailsPanel_2");

	}
	public  Credentials googleOathCredentials(String loginUserName,String selectionType) throws IOException {

		try {

			String TOKENS_DIRECTORY_PATH = null;
			final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();	
			if (System.getProperty("os.name").toLowerCase().contains("windows")) {

				TOKENS_DIRECTORY_PATH = System.getenv("APPDATA") + File.separator + ToolDetails.messageboxtitle
						+ File.separator +selectionType+File.separator+ loginUserName;

			} else {

				TOKENS_DIRECTORY_PATH = System.getProperty("user.home") + File.separator + "Library" + File.separator
						+ "Application Support" + File.separator + ToolDetails.messageboxtitle + File.separator+selectionType+File.separator+
						loginUserName;
			}

			File checkClient = new File(TOKENS_DIRECTORY_PATH);
			if (!checkClient.exists()) {
				JOptionPane.showMessageDialog(null, "Redirect to Browser please Allow Access", ToolDetails.messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE);
			}

			List<String> scope = new ArrayList<String>();
			scope.add("https://www.googleapis.com/auth/drive");
			scope.add("https://www.googleapis.com/auth/calendar");
			scope.add("https://www.googleapis.com/auth/contacts");
			scope.add("https://mail.google.com");
			scope.add("https://www.googleapis.com/auth/gmail.labels");
			scope.add("https://www.googleapis.com/auth/photoslibrary");
			

			InputStream in = EmailWizardApplication.class.getResourceAsStream("/credentials.json");
			GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

			FileDataStoreFactory dataStoreFactory = new FileDataStoreFactory(new File(TOKENS_DIRECTORY_PATH));

			String clientfileName=selectionType+loginUserName; 
			DataStore<StoredCredential> dataStore = dataStoreFactory.getDataStore(getMd5(clientfileName).substring(1,10));

			GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
					clientSecrets, scope).setCredentialDataStore(dataStore).setAccessType("offline").build();

			LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(9999).build();

			oathCredential = new AuthorizationCodeInstalledApp(flow, receiver).authorize(loginUserName);

			clientId = clientSecrets.getDetails().getClientId();
			clientSecret = clientSecrets.getDetails().getClientSecret();


		} catch (Exception e) {
			// TODO: handle exception

			e.printStackTrace();
	
		}
		
			CardLayout card = (CardLayout) EmailWizardApplication.CardLayout.getLayout();
			card.show(EmailWizardApplication.CardLayout, "GoogleDownloadOptions_3");
			
			
			oathCredentials= UserCredentials.newBuilder().setClientId(clientId).setClientSecret(clientSecret)
					.setRefreshToken(oathCredential.getRefreshToken()).build();

		return oathCredentials;


	}
	

	
public void setGoogleCredential(GoogleCredential credential)
{
	this.credential=credential;
}
public GoogleCredential getGoogleCredentials()
{
	return credential;
	
}

public void setoathGoogleCredential(Credential oathCredential)
{
	this.oathCredential=oathCredential;
}
public Credential getoathGoogleCredential()
{
	return oathCredential;
	
}
public Credentials getoathGoogleCredentials()
{
	return oathCredentials;
	
}
	
public static String getMd5(String input)
{
    try {

        // Static getInstance method is called with hashing MD5
        MessageDigest md = MessageDigest.getInstance("MD5");

        // digest() method is called to calculate message digest
        // of an input digest() return array of byte
        byte[] messageDigest = md.digest(input.getBytes());

        // Convert byte array into signum representation
        BigInteger no = new BigInteger(1, messageDigest);

        // Convert message digest into hex value
        String hashtext = no.toString(16);
        while (hashtext.length() < 32) {
            hashtext = "0" + hashtext;
        }
        return hashtext;
    }

    // For specifying wrong message digest algorithms
    catch (NoSuchAlgorithmException e) {
        throw new RuntimeException(e);
    }
}
}
