package com.api.ews;
import java.awt.Desktop;
import java.awt.Desktop.Action;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.URLConnection;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.SimpleFormatter;

import org.json.simple.JSONObject;
import org.slf4j.LoggerFactory;

import com.aspose.email.ITokenProvider;
import com.chilkatsoft.CkGlobal;
import com.chilkatsoft.CkJsonObject;
import com.chilkatsoft.CkOAuth2;
import com.google.api.client.util.Preconditions;
import com.main.EmailWizardApplication;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ConnectingIdType;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.misc.ITraceListener;
import microsoft.exchange.webservices.data.misc.ImpersonatedUserId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FolderView;


public class EWSAccesToken {
	
	public  org.slf4j.Logger logger=LoggerFactory.getLogger(EmailWizardApplication.class);
	String refreshToken;
	
	static {

		FileHandler fh;
		try {
			
			fh = new FileHandler(System.getProperty("java.io.tmpdir") + File.separator + "chilkat.log");
			SimpleFormatter formatter = new SimpleFormatter();
			fh.setFormatter(formatter);


		} catch (SecurityException e) {

			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		try {
			InputStream in = EWSAccesToken.class.getResourceAsStream("/chilkat.dll");
			byte[] buffer = new byte[1024];
			int read = -1;

			File temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkat.dll");
			int i = 0;
			while (temp.exists()) {

				temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkat" + i + ".dll");
				i++;
			}
			FileOutputStream fos = null;
			try {

				fos = new FileOutputStream(temp);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
				// logger.warning(e.getMessage());
			}
			try {
				while ((read = in.read(buffer)) != -1) {
					fos.write(buffer, 0, read);
				}
			} catch (IOException e) {
				// logger.warning(e.getMessage());
			}
			try {
				fos.close();
			} catch (IOException e) {
				// logger.warning(e.getMessage());
			}
			try {
				in.close();
			} catch (IOException e) {
				
				// logger.warning(e.getMessage());
			}

			System.load(temp.getAbsolutePath());
		
			
		} catch (UnsatisfiedLinkError | Exception e) {
			e.printStackTrace();
	
		}

	
	}

	public String getAccessToken() {

		CkGlobal glob = new CkGlobal();
		boolean success = glob.UnlockBundle("SNKRWT.CB1122022_Q5uG5AzJlRm1");
		if (success != true) {
			
			return null;
		}

		int status = glob.get_UnlockStatus();
		if (status == 2) {
		
		} else {
			
		}
	
		String accountId=null;	
	    CkOAuth2 oauth2 = new CkOAuth2();
	    boolean success1;


	    oauth2.put_ListenPort(3017);

	    oauth2.put_AuthorizationEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");
	    oauth2.put_TokenEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/token");	    
	    oauth2.put_ClientId("1994b9da-99f0-4c6b-9f78-75c47976d339");	 
	    oauth2.put_ClientSecret("Ub4.5yQl-243bYQyfLx.W4MM-w3rfMom8-");
	    oauth2.put_CodeChallenge(false);
	    
	    oauth2.put_Scope("openid profile offline_access https://outlook.office365.com/SMTP.Send https://outlook.office365.com/POP.AccessAsUser.All https://outlook.office365.com/IMAP.AccessAsUser.All https://outlook.office365.com/EWS.AccessAsUser.All");	    
	   
	   oauth2.put_RedirectAllowHtml("<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #f1f9ff; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.sysinfotools.com/sysinfo-img/logo.png' alt='logo'><p style='font-size: 27px;color: #19a300;'>Authentication Successful</p><p>Please, go back to the SysInfo OneDrive Migration Wizard!</p></div></div>");
	    oauth2.put_RedirectDenyHtml("<div style='text-align: center; display: flex; justify-content: center!important; align-items: center; height: 100%;'><div style='flex: 0 0 33.333333%; max-width: 33.333333%; position: relative; width: 100%; background: #f1f9ff; border-radius: 20px; padding: 1rem; border: 1px solid #ddd;'><img src='https://www.sysinfotools.com/sysinfo-img/logo.png' alt='logo'><p style='font-size: 27px;color: #ff0000;'>Access Denied</p><p>Please, go back to the Email Backup !</p></div></div>");
	    
	    String url = oauth2.startAuth();
	    if (oauth2.get_LastMethodSuccess() != true) {
	        System.out.println(oauth2.lastErrorText());
	        
	        }

	    try {
	    	URI uri = new URI(url);
	    	URL urls=uri.toURL();	    	
	        
	        URLConnection   conn = (HttpURLConnection) urls.openConnection();
			conn.setUseCaches(false);
			conn.setDefaultUseCaches(false);
		     
	        browse(conn.getURL().toURI().toString());
			//Desktop.getDesktop().browse(conn.getURL().toURI());
		} catch (IOException e) {
			
			e.printStackTrace();
		} catch (URISyntaxException e) {
			
			e.printStackTrace();
		}
	    int numMsWaited = 0;
	    while ((numMsWaited < 300000) && (oauth2.get_AuthFlowState() < 3)) {
	        oauth2.SleepMs(100);
	        numMsWaited = numMsWaited+100;
	        }


	    if (oauth2.get_AuthFlowState() < 3) {
	        oauth2.Cancel();
	      
	        }

	
	    if (oauth2.get_AuthFlowState() == 5) {
	        System.out.println("OAuth2 failed to complete.");
	        
	        }

	    if (oauth2.get_AuthFlowState() == 4) {
	        System.out.println("OAuth2 authorization was denied.");
	     
	        
	        }

	    if (oauth2.get_AuthFlowState() != 3) {
	        System.out.println("Unexpected AuthFlowState:" + oauth2.get_AuthFlowState());
	        
	        }
	    
	    
	 // Get the full JSON response:
	    CkJsonObject json = new CkJsonObject();
	    json.Load(oauth2.accessTokenResponse());
	    json.put_EmitCompact(false);
	    String iValue = json.stringAt(4);
	    String refreshToken = json.stringAt(5);
	    setRefreshToken(refreshToken);
	    System.out.println(json.emit());
	   
	    return iValue;
	   
	}
	
	public String getRefreshAccessToken(String refresh) {

	    CkOAuth2 oauth2 = new CkOAuth2();
	    oauth2.put_TokenEndpoint("https://login.microsoftonline.com/common/oauth2/v2.0/token");

	    // Replace these with actual values.
	    oauth2.put_ClientId("1994b9da-99f0-4c6b-9f78-75c47976d339");
	    oauth2.put_ClientSecret("Ub4.5yQl-243bYQyfLx.W4MM-w3rfMom8-");

	    // Get the "refresh_token"
	    oauth2.put_RefreshToken(refresh);

	    // Send the HTTP POST to refresh the access token..
	   boolean  success = oauth2.RefreshAccessToken();
	    if (success != true) {
	     //   System.out.println(oauth2.lastErrorText());
	        return null;
	        }
	    
		logger.info("OAuth2 authorization granted!");
	    

	    //System.out.println("New refresh token: " + oauth2.refreshToken());	   
	    //System.out.println("New Access Token = " + oauth2.accessToken());
	    
	 // Get the full JSON response:
	    CkJsonObject json = new CkJsonObject();
	    json.Load(oauth2.accessTokenResponse());
	    json.put_EmitCompact(false);
	    String iValue = json.stringAt(4);
	  //  System.out.println(iValue);
	    String refreshToken = json.stringAt(5);
	    setRefreshToken(refreshToken);
	   // System.out.println(json.emit());
	    return iValue;
	  
	   
	}
	public static void browse(String url) {
	    Preconditions.checkNotNull(url);
	    // Ask user to open in their browser using copy-paste
	    System.out.println("Please open the following address in your browser:");
	   // System.out.println("  " + url);
	    // Attempt to open it in the browser
	    try {
	      if (Desktop.isDesktopSupported()) {
	        Desktop desktop = Desktop.getDesktop();
	        if (desktop.isSupported(Action.BROWSE)) {
	          System.out.println("Attempting to open that address in the default browser now...");
	          desktop.browse(URI.create(url));
	        }
	      }
	    } catch (IOException e) {
	     // LOGGER.log(Level.WARNING, "Unable to open browser", e);
	    } catch (InternalError e) {
	      // A bug in a JRE can cause Desktop.isDesktopSupported() to throw an
	      // InternalError rather than returning false. The error reads,
	      // "Can't connect to X11 window server using ':0.0' as the value of the
	      // DISPLAY variable." The exact error message may vary slightly.
	     // LOGGER.log(Level.WARNING, "Unable to open browser", e);
	    }
	  }
	
	public String getRefreshToken()
	{
		return refreshToken;
	}
	
	public void setRefreshToken(String refreshToken)
	{
		this.refreshToken=refreshToken;
	}

}
