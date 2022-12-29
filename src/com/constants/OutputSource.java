package com.constants;

import java.util.ArrayList;
import java.util.Collections;

public enum OutputSource {
		
	 PST,
	 OST,
	 MSG,
	 EML,
	 EMLX,
	 HTML,
	 PDF,
	 Office365,
	 GMAIL,
	 AWS,
	 YAHOO,
	 AOL,
	 IMAP,
	 MBOX,
	 CSV,
	 ICLOUD,
	 GODADDY,
	 ZOHO_EMAIL,
	 HOTMAIL,
	 YANDEX,
	 GSUITE,
	 RTF,
	 XPS,
	 EMF,
	 DOCX,
	 JPEG,
	 DOCM,
	 TEXT,
	 PNG,
	 TIFF,
	 SVG,
	 EPUB,
	 DOTM,
	 BMP,
	 GIF,
	 OTT,
	 WORLD_ML,
	 ODT, 
	 MS_Office_365,
	 HostGator,
	 GMAIL_APP,
	 Hotmail;
	
	public static ArrayList<String> imapClientOutputFormat=new ArrayList<>();
    
	static {		
		imapClientOutputFormat.add(GMAIL.name());
		imapClientOutputFormat.add(AWS.name());
		imapClientOutputFormat.add(Hotmail.name());
		imapClientOutputFormat.add(Office365.name());
		imapClientOutputFormat.add(YAHOO.name());
		imapClientOutputFormat.add(AOL.name());
		imapClientOutputFormat.add(IMAP.name());
		imapClientOutputFormat.add(HostGator.name());
		imapClientOutputFormat.add(ICLOUD.name());
		imapClientOutputFormat.add(GODADDY.name());
		imapClientOutputFormat.add(ZOHO_EMAIL.name());
		imapClientOutputFormat.add(YANDEX.name());
		imapClientOutputFormat.add(MS_Office_365.name());
		imapClientOutputFormat.add(GMAIL_APP.name());
		imapClientOutputFormat.add(GSUITE.name());
		
        }
    }



