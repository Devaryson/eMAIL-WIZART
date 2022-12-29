package com.util;

import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

import com.aspose.email.MailMessage;
import com.main.EmailWizardApplication;

public class RemoveDuplicateEmails {
	
	String subject;
	String from;
	String to;
	String Cc;
	String body;
	String date;
	MailMessage msg;
	private static RemoveDuplicateEmails removeDuplicate = null;
	public void setMailMessage(MailMessage msg) {
		this.msg=msg;		
		setSubject(msg.getSubject());
		setFrom(msg.getFrom().getAddress());
		setDate(msg.getDate().toString());		
		setBody(msg.getBody());
	
	}
	public static RemoveDuplicateEmails getInstance()
    {
        if (removeDuplicate == null)
        	removeDuplicate = new RemoveDuplicateEmails();
  
        return removeDuplicate;
    }
	
	public String getHashString()
	{
		StringBuilder hashString=new StringBuilder();
		
		if(EmailWizardApplication.chckbxSkip_body.isSelected())
		{
			hashString.append(getBody());
		}
		if(EmailWizardApplication.chckbxSkip_subject.isSelected())
		{
			hashString.append(getSubject());
		}
		if(EmailWizardApplication.chckbxSkip_date.isSelected())
		{
			hashString.append(getDate());
		}
		
		if(EmailWizardApplication.chckbxSkip_from.isSelected())
		{
			hashString.append(getFrom());
		}

	return getMd5(hashString.toString());
		
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
            System.out.println(hashtext);
            return hashtext;
        }
 
        // For specifying wrong message digest algorithms
        catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        }
    }
	public String getSubject() {
		return subject;
	}
	public void setSubject(String subject) {
		this.subject = subject;
	}
	public String getFrom() {
		return from;
	}
	public void setFrom(String from) {
		this.from = from;
	}
	public String getTo() {
		return to;
	}
	public void setTo(String bcc) {
		to = to;
	}
	
	public String getBody() {
		return body;
	}
	public void setBody(String body) {
		this.body = body;
	}
	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}

}
