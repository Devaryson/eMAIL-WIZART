package com.downoad.googleapp;

import java.awt.CardLayout;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.MalformedURLException;
import java.net.NoRouteToHostException;
import java.net.URI;
import java.net.URL;
import java.net.URLConnection;
import java.security.GeneralSecurityException;
import java.security.Principal;
import java.security.PrivateKey;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.swing.ImageIcon;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;

import org.threeten.bp.LocalDate;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.auth.oauth2.StoredCredential;
import com.google.api.client.extensions.android.util.store.FileDataStoreFactory;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.DataStore;
import com.google.api.gax.core.CredentialsProvider;
import com.google.api.gax.core.FixedCredentialsProvider;
import com.google.api.services.drive.Drive;
import com.google.auth.Credentials;
import com.google.auth.oauth2.AccessToken;
import com.google.auth.oauth2.GoogleCredentials;
import com.google.auth.oauth2.ServiceAccountCredentials;
import com.google.auth.oauth2.UserCredentials;
import com.google.photos.library.v1.PhotosLibraryClient;
import com.google.photos.library.v1.PhotosLibrarySettings;
import com.google.photos.library.v1.internal.InternalPhotosLibraryClient.ListAlbumsPagedResponse;
import com.google.photos.library.v1.internal.InternalPhotosLibraryClient.ListMediaItemsPagedResponse;
import com.google.photos.library.v1.internal.InternalPhotosLibraryClient.SearchMediaItemsPagedResponse;
import com.google.photos.types.proto.Album;
import com.google.photos.types.proto.MediaItem;
import com.google.photos.types.proto.MediaMetadata;
import com.google.photos.types.proto.Photo;
import com.main.EmailWizardApplication;

import io.grpc.Context.Storage;

public class PhotosBackup {

	static CardLayout card;
	public static Credentials auth;
	public static PhotosLibraryClient photosLibraryClient;
	final static JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	static GoogleCredential credentials;
	public static Drive driveService = null;

	static String serviceAccountUsers;

	@SuppressWarnings("deprecation")
	public void googleCredentials(String serviceAccountId, String serviceAccountUser, String p12File)
			throws GeneralSecurityException, IOException {

		serviceAccountUsers = serviceAccountUser;
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

		com.google.api.client.googleapis.auth.oauth2.GoogleCredential.Builder credentialBuilder = new GoogleCredential.Builder()
				.setTransport(HTTP_TRANSPORT).setJsonFactory(JSON_FACTORY).setServiceAccountId(serviceAccountId)
				.setServiceAccountScopes(Collections.singleton("https://www.googleapis.com/auth/photoslibrary"))
				.setServiceAccountUser(serviceAccountUser).setServiceAccountPrivateKeyFromP12File(new File(p12File));

		GoogleCredential gc = credentialBuilder.build();

		GoogleCredentials sac = ServiceAccountCredentials.newBuilder()

				.setPrivateKey(gc.getServiceAccountPrivateKey()).setPrivateKeyId(gc.getServiceAccountPrivateKeyId())
				// .setServiceAccountUser(serviceAccountUser)
				.setScopes(gc.getServiceAccountScopes())
				// .setAccessToken(new AccessToken(gc.getAccessToken(), calendar.getTime()))
				.build();

		// Latest generation Google libs, GoogleCredentials extends Credentials
		CredentialsProvider cp = FixedCredentialsProvider.create(sac);

		PhotosLibrarySettings settings = PhotosLibrarySettings.newBuilder().setCredentialsProvider(cp).build();

		photosLibraryClient = PhotosLibraryClient.initialize(settings);

		photoDownload(photosLibraryClient, new File(EmailWizardApplication.detinationPath));
	}

	public void googleCredentials(Credentials credentials) throws GeneralSecurityException, IOException {

		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		PhotosLibrarySettings settings = PhotosLibrarySettings.newBuilder()
				.setCredentialsProvider(FixedCredentialsProvider.create(credentials)).build();
		photosLibraryClient = PhotosLibraryClient.initialize(settings);
		photoDownload(photosLibraryClient, new File(EmailWizardApplication.detinationPath));
	}

	public static void albumDownload(PhotosLibraryClient photosLibraryClient, File folderName) throws IOException {
		ListAlbumsPagedResponse response = photosLibraryClient.listAlbums();
		for (Album album : response.iterateAll()) {

			File albumName = new File(folderName.getAbsolutePath() + File.separator + album.getTitle());
			albumName.mkdirs();
			SearchMediaItemsPagedResponse responseMedia = photosLibraryClient.searchMediaItems(album.getId());

			for (MediaItem item : responseMedia.iterateAll()) {
				googlePhotosDownloading(item, albumName);

			}

		}
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
	

	public static void photoDownload(PhotosLibraryClient photosLibraryClient, File foldername) {
		try {
			EmailWizardApplication.progressBar_Downloading.setValue(0);
			EmailWizardApplication.progressBar_Downloading.setVisible(false);
			EmailWizardApplication.progressBar_Downloading.setMaximum(100);
			EmailWizardApplication.progressBar_Downloading.setStringPainted(true);
			EmailWizardApplication.lblDownloading.setVisible(true);

			DefaultTableModel model = (DefaultTableModel) EmailWizardApplication.table_Downloading.getModel();
			ListMediaItemsPagedResponse responseMedia = photosLibraryClient.listMediaItems();
			int i = 0;

			Iterable<MediaItem> mediaItem = responseMedia.iterateAll();
			MediaItem itemTemp = null;
			for (MediaItem item : responseMedia.iterateAll()) {

				try {
					checkStopAndDemo(i);
					itemTemp = item;
					googlePhotosDownloading(item, foldername);

					MediaMetadata metadata = item.getMediaMetadata();
					String mediaType = null;
					if (metadata.hasVideo()) {
						mediaType = "Video Downloaded";
					} else if (metadata.hasPhoto()) {
						mediaType = "Photo downloaded";
					}

					System.out.println("photo downloaded");
					
					
					EmailWizardApplication.downloadingFileName.setText(validFileName(item.getFilename()));			
					EmailWizardApplication.modelDownloading.setValueAt(i, EmailWizardApplication.rownCount,0);	
					i++;
				} catch (Exception e) {
					if (e.getMessage().contains("www.googleapis.com")
							|| e.getMessage().contains("oauth2.googleapis.com")
							|| e.getMessage().contains("No route to host: connect")) {

						EmailWizardApplication.lblNoInternetConnection.setVisible(true);
						while (!checkInternet()) {
							System.out.println("No Internet!!");
						}

						EmailWizardApplication.lblNoInternetConnection.setVisible(false);
					}
				}

			}

			ListAlbumsPagedResponse response = photosLibraryClient.listAlbums();

			for (Album album : response.iterateAll()) {

				File albumName = new File(
						foldername.getAbsolutePath() + File.separator + validFileName(album.getTitle()));
				albumName.mkdirs();
				SearchMediaItemsPagedResponse responseMediaAlbum = photosLibraryClient.searchMediaItems(album.getId());

				MediaItem itemTempAlbum = null;
				for (MediaItem item : responseMediaAlbum.iterateAll()) {

					try {

						checkStopAndDemo(i);
						itemTempAlbum = item;
						googlePhotosDownloading(item, albumName);

						System.out.println("ablum downloaded");
					} catch (Exception e) {
						if (e.getMessage().contains("www.googleapis.com")
								|| e.getMessage().contains("oauth2.googleapis.com")
								|| e.getMessage().contains("No route to host: connect")) {

							EmailWizardApplication.lblNoInternetConnection.setVisible(true);
							while (!checkInternet()) {
								System.out.println("No Internet!!");
							}

							EmailWizardApplication.lblNoInternetConnection.setVisible(false);
						}
					}

				}

				EmailWizardApplication.downloadingFileName.setText(validFileName(album.getTitle()));			
				EmailWizardApplication.modelDownloading.setValueAt(i, EmailWizardApplication.rownCount, 0);
				i++;

			}
		} catch (Exception e) {

			System.out.println(e.getMessage());
		}

		EmailWizardApplication.lblDownloading.setVisible(false);


	}

	public static void googlePhotosDownloading(MediaItem item, File foldername) throws IOException {

		OutputStream os = null;
		InputStream is = null;

		MediaMetadata metadata = item.getMediaMetadata();
		String fileUrl = null;
		if (metadata.hasPhoto()) {
			fileUrl = item.getBaseUrl() + "=d";
		} else if (metadata.hasVideo()) {
			fileUrl = item.getBaseUrl() + "=dv";
		}

		String outputPath = foldername.getAbsolutePath() + File.separator + validFileName(item.getFilename());

		// create a url object
		URL url = new URL(fileUrl);
		// connection to the file
		URLConnection connection = url.openConnection();
		// get input stream to the file
		is = connection.getInputStream();
		// get output stream to download file
		os = new FileOutputStream(outputPath);
		final byte[] b = new byte[2048];
		int length;
		// read from input stream and write to output stream
		while ((length = is.read(b)) != -1) {
			os.write(b, 0, length);
		}

		os.close();

	}

	public static String validFileName(String fileName) {
		// System.out.println(fileName.replaceAll("[^a-zA-Z0-9\\.\\-]", "_"));
		return fileName.replaceAll("[^a-zA-Z0-9\\.\\-]", "_").trim();

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

}
