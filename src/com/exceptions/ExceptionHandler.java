package com.exceptions;

import javax.swing.ImageIcon;
import javax.swing.JOptionPane;

import com.main.EmailWizardApplication;
import com.tool.info.ToolDetails;

public class ExceptionHandler  {
	
	Exception exception;
	EmailWizardApplication mf;
	
	public ExceptionHandler(Exception exception,EmailWizardApplication mf)
	{
		
		this.exception=exception;
		this.mf=mf;
		
	}
	public ExceptionHandler(Exception exception)
	{
		
		this.exception=exception;
		
	}
public void loginExceptionHandler()
{
	
	if (exception.getMessage()
			.equalsIgnoreCase("AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
		JOptionPane.showMessageDialog(mf,
				"Connection Not Estalished with Imap server please check your Credantial Otherwise allow 3rd party app to acess your account",
				ToolDetails.messageboxtitle, JOptionPane.ERROR_MESSAGE,
				new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
	} else if (exception.getMessage().contains("Application-specific password required:")) {
		JOptionPane.showMessageDialog(mf, "Application specific password required",
				ToolDetails.messageboxtitle, JOptionPane.ERROR_MESSAGE,
				new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
	} else if (exception.getMessage().contains("Unable connect to the server.")) {
		JOptionPane.showMessageDialog(mf, "Unable connect to the server.",
				ToolDetails.messageboxtitle, JOptionPane.ERROR_MESSAGE,
				new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
	}
	 else if (exception.getMessage().contains("The request failed. outlook.office365.com")) {
			JOptionPane.showMessageDialog(mf, "Unable connect to the server please check you internet Connction.",
					ToolDetails.messageboxtitle, JOptionPane.ERROR_MESSAGE,
					new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
		}

	else {
		JOptionPane.showMessageDialog(mf,
				"Connection not established, Please enter valid email address and its passowrd",
				ToolDetails.messageboxtitle, JOptionPane.ERROR_MESSAGE,
				new ImageIcon(EmailWizardApplication.class.getResource("/information.png")));
	}
	
}
public boolean migrationExceptionHandler()
{
	       if (exception.getMessage().contains("No connection could be made because the target machine actively refused it.")
			|| exception.getMessage().contains("ConnectFailure") 
			|| exception.getMessage().contains("Rate limit hit")
			|| exception.getMessage().contains("imap.gmail.com")
			|| exception.getMessage().contains("Operation failed") 
			|| exception.getMessage().contains("Remote host terminated the handshake") 			
			|| exception.getMessage().contains("Operation has been canceled")
			|| exception.getMessage().contains("No route to host: connect")					
			|| exception.getMessage().contains("ImapException: java.net.UnknownHostException: outlook.office365.com")
			|| exception.getMessage().contains("outlook.office365.com")
			|| exception.getMessage().contains("No connection could be made because the target machine actively refused it")
			|| exception.getMessage().contains("Server logging out") 
			|| exception.getMessage().contains("java.net.UnknownHostException: imap.gmail.com") 
			|| exception.getMessage().contains("Object has been disposed.")
			|| exception.getMessage().contains("* BYE Server shutting down.")
			|| exception.getMessage().contains("An error has arisen while command is sent")) {

	     	return true;
	
	}
	return false;
}
public boolean appendExceptionHandler()
{
	       if (exception.getMessage().contains("AE_858_3_0006 NO [SERVERBUG] APPEND Server error - Please try again later")
	        || exception.getMessage().contains("AE_1_2_0536 BAD Command Argument Error.")
			|| exception.getMessage().contains("The operation 'AppendMessage' terminated. Timeout '60000' has been reached.")
			|| exception.getMessage().contains("The operation 'AppendMessage' terminated. Timeout '59997' has been reached.")
			|| exception.getMessage().contains("The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {

	     	return true;
	
	}
	return false;
}
public boolean GsuiteExceptionHandler()
{
	if (exception.getMessage().contains("www.googleapis.com")
		|| exception.getMessage().contains("oauth2.googleapis.com")
		|| exception.getMessage().contains("No route to host: connect")
		|| exception.getMessage().contains("gmail.googleapis.com")
		|| exception.getMessage().contains("Failed to refresh access token: Connection reset")
		|| exception.getMessage().contains("Connection reset")
		|| exception.getMessage().contains("Software caused connection abort: connect")
		|| exception.getMessage().contains("imap.gmail.com")
		|| exception.getMessage().contains("Read timed out")) {
		return true;
	
}
	return false;
}
	

}
