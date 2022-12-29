package com.constants;

import java.util.Arrays;
import java.util.List;

import javax.swing.DefaultListModel;

public enum InputSource {
	
   	 
	 GMAIL("GMAIL"),	 
	 GMAIL_APP("GMAIL_APP"),
	 GSUITE("GSUITE"),
	 Office365("Office365"),
	 MS_Office_365("MS_Office_365"),
	 HOTMAIL("HOTMAIL"),
	 YAHOO("YAHOO"),
	 AOL("AOL"),
	 AWS("AWS"),
	 GODADDY("GODADDY"),	
	 IMAP("IMAP"),
	 HOSTGATOR("HOSTGATOR"),
	 ICLOUD("ICLOUD"),
	 ZOHO_EMAIL("ZOHO_EMAIL"),	
	 YANDEX("YANDEX"),		  	
	 AIM("AIM"),
	 ARCOR("ARCOR"),
	 ARUBA("ARUBA"),
	 ASIA_COM("ASIA.COM"),
	 AT_AND_T("AT&T"),
	 AXIGEN("AXIGEN"),
	 ONE_AND_ONE_MAIL("1&1 Mail"),
	 ONE_TWO_SIX("126 Mail"),
	 ONE_SIX_THREE("163 Mail"),
	 Bulk("Bulk");

	private String value; 
	
	private InputSource(String string){  
	this.value=string;  
	
	}
	
	public String getValue(){
		
		return value;
	}
	
	public static DefaultListModel<InputSource> getDefaultListModel() {

		DefaultListModel<InputSource> defaultListModel = new DefaultListModel<InputSource>();
		List<InputSource> enumList = Arrays.asList(InputSource.class.getEnumConstants());
		for (int i = 0; i < enumList.size(); i++) {
			defaultListModel.add(i, enumList.get(i));

		}
		return defaultListModel;

	}
   

    }


