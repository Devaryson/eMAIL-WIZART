package com.util;

import java.text.DecimalFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.aspose.email.HeaderCollection;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapMessageInfo;
import com.aspose.email.ImapNamespace;
import com.aspose.email.MailMessage;
import com.google.api.services.gmail.model.MessagePartHeader;
import com.main.EmailWizardApplication;

public interface FileNamingUtils {
	
	public static String changeImapFolderNameDelimeter(ImapClient clientforimap_input,String folderName)
	{
		
		ImapNamespace[] s=clientforimap_input.getNamespaces();
		for (ImapNamespace imapNamespace : s) {	
			//System.out.println(imapNamespace.getHierarchyDelimiter());
			if(!imapNamespace.getHierarchyDelimiter().equals("/"))
			{
				String removeSlash=folderName.replaceAll("/", "%");				
				return removeSlash.replace(imapNamespace.getHierarchyDelimiter(), "/");

			}
			
		}
		return folderName;
	}
	public static String changeBackImapFolderNameDelimeter(ImapClient clientforimap_input,String folderName)
	{
		ImapNamespace[] s=clientforimap_input.getNamespaces();
		for (ImapNamespace imapNamespace : s) {	
			if(imapNamespace.getHierarchyDelimiter()!="/"&&imapNamespace.getHierarchyDelimiter()=="|")
			{				
				String addSlash=folderName.replace("/", "|");				
				return addSlash.replaceAll("%s", "/");

			}
			
		}
		return folderName;
	}
	

	public static String buildFileName(MailMessage mailMessage, Long timeStamp) {

		String subjectName;
		String from;
		String filename = null;

		try {
			subjectName = namingConvention(mailMessage.getSubject());
		} catch (Exception e) {

			subjectName = "";
		}
		if (subjectName.length() > 40) {
			subjectName = subjectName.substring(0, 40);
		}

		try {
			from = namingConvention(mailMessage.getFrom().getAddress());
		} catch (Exception e) {

			from = "";
		}
		if (from.length() > 20) {
			from = from.substring(0, 20);
		}
		return buildFileName( subjectName, from, filename,  timeStamp);
	}
	public static String buildFileName(ImapMessageInfo messagePartHeaderlist, Long timeStamp) {

		String subjectName;
		String from;
		String filename = null;

		try {
	
			subjectName = namingConvention(messagePartHeaderlist.getSubject());
		} catch (NoSuchElementException e) {

			subjectName = "";
		}
		if (subjectName.length() > 40) {
			subjectName = subjectName.substring(0, 40);
		}

		try {

			from = namingConvention(messagePartHeaderlist.getFrom().getAddress());
		} catch (NoSuchElementException e) {

			from = "";
		}
		if (from.length() > 20) {
			from = from.substring(0, 20);
		}
		return buildFileName( subjectName, from, filename,  timeStamp);
	}

		public static String buildFileName(List<MessagePartHeader> messagePartHeaderlist,Long timeStamp) {

			String subjectName;
			String from;
			String filename = null;
			
			try {
				MessagePartHeader messagePartHeader = messagePartHeaderlist.stream().filter(x -> x.getName().equals("Subject")).findFirst().get();
				subjectName = namingConvention(messagePartHeader.getValue());
			} catch (NoSuchElementException e) {
				
				subjectName="";
			}
			if (subjectName.length() > 40) {
				subjectName = subjectName.substring(0, 40);
			}
			
			try {
				MessagePartHeader messagePartHeader = messagePartHeaderlist.stream().filter(x -> x.getName().equals("From")).findFirst().get();
				from = namingConvention(messagePartHeader.getValue());
			} catch (NoSuchElementException e) {
				
				from="";
			}
			if (from.length() > 20) {
				from = from.substring(0, 20);
			}
			
			return buildFileName(subjectName, from, filename,  timeStamp);
		
	}
	
	public static String buildFileName(HeaderCollection messagePartHeaderlist, Long timeStamp) {
		
		String subjectName;
		String from;
		String filename = null;

		try {
			// MessagePartHeader messagePartHeader = messagePartHeaderlist.stream().filter(x
			// -> x.getName().equals("Subject")).findFirst().get();
			subjectName = namingConvention(messagePartHeaderlist.get("Subject"));
		} catch (NoSuchElementException e) {

			subjectName = "";
		}
		if (subjectName.length() > 40) {
			subjectName = subjectName.substring(0, 40);
		}

		try {
			// MessagePartHeader messagePartHeader = messagePartHeaderlist.stream().filter(x
			// -> x.getName().equals("From")).findFirst().get();
			from = namingConvention(messagePartHeaderlist.get("From"));
		} catch (NoSuchElementException e) {

			from = "";
		}
		if (from.length() > 20) {
			from = from.substring(0, 20);
		}
		return buildFileName(subjectName, from, filename,  timeStamp);
		
	}
	
	public static String buildFileName(String subjectName,String from,String filename, Long timeStamp)
	{
		String dstr = "";
		Date d;
		String combox_selected = EmailWizardApplication.comboBoxNamingConvention.getSelectedItem().toString();
		try {

			Calendar cal = Calendar.getInstance();
			cal.setTimeInMillis(timeStamp);

			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DAY_OF_MONTH);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = subjectName;
		} 
		if (combox_selected.contains("Subject_Date")) {
			filename = subjectName + "_" + dstr;
		} 
		if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + subjectName;
		} 
		if (combox_selected.contains("From_Subject_Date")) {
			filename = from + "_" + subjectName + "_" + dstr;
		} 
		if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + from + "_" + subjectName;
		}
		
		return getRidOfIllegalFileNameCharacters(filename);
	}


	public static String namingConvention(String subject) {

		String subjectName = null;
		if (subject != null) {
			subjectName = subject;
			if (subjectName.length() > 40) {
				subjectName = subject.substring(0, 40);
			}

		} else {
			subjectName = "";
		}

		return getRidOfIllegalFileNameCharacters(subjectName).trim();

	}

	public static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
				.replace("//s", "").replace("\"", "");
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName.trim();
	}
	public static String validFileNameForWindows(String strName) {
		String strLegalName = strName.replace(":", "-").replace("?", "-")
				.replace("|", "-").replace("*", "-").replace("<", "-").replace(">", "-").replace("\t", "-")
				.replace("//s", "-");
		 return strLegalName.trim();
	}
	public static String buildImapFolderName(String folderName) {
		return folderName.trim();
		
	}
	public static String buildImapFolderNameLength(String folderName) {
		if (folderName.length() > 40) {
			folderName = folderName.substring(0, 30);
		}
		return folderName.trim();
		
	}
	public static boolean isValidName(String text)
	{
	    Pattern pattern = Pattern.compile(
	        "# Match a valid Windows filename (unspecified file system).          \n" +
	        "^                                # Anchor to start of string.        \n" +
	        "(?!                              # Assert filename is not: CON, PRN, \n" +
	        "  (?:                            # AUX, NUL, COM1, COM2, COM3, COM4, \n" +
	        "    CON|PRN|AUX|NUL|             # COM5, COM6, COM7, COM8, COM9,     \n" +
	        "    COM[1-9]|LPT[1-9]            # LPT1, LPT2, LPT3, LPT4, LPT5,     \n" +
	        "  )                              # LPT6, LPT7, LPT8, and LPT9...     \n" +
	        "  (?:\\.[^.]*)?                  # followed by optional extension    \n" +
	        "  $                              # and end of string                 \n" +
	        ")                                # End negative lookahead assertion. \n" +
	        "[^<>:\"/\\\\|?*\\x00-\\x1F]*     # Zero or more valid filename chars.\n" +
	        "[^<>:\"/\\\\|?*\\x00-\\x1F\\ .]  # Last char is not a space or dot.  \n" +
	        "$                                # Anchor to end of string.            ", 
	        Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE | Pattern.COMMENTS);
	    Matcher matcher = pattern.matcher(text);
	    boolean isMatch = matcher.matches();
	    return isMatch;
	}
	

}
