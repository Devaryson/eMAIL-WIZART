package com.api.ews;

import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

import javax.swing.table.DefaultTableModel;
import javax.swing.tree.DefaultMutableTreeNode;

import com.aspose.email.microsoft.schemas.exchange.services._2006.types.EmailAddress;
import com.main.EmailWizardApplication;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ConnectingIdType;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.FolderTraversal;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.TokenCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.credential.WebProxyCredentials;
import microsoft.exchange.webservices.data.misc.ImpersonatedUserId;
import microsoft.exchange.webservices.data.property.complex.EmailAddressCollection;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FolderView;

public class EWSOffice{
	
	private static String username;
	private static String password;
	private ExchangeService service ;
	private final static String HOST="https://outlook.office365.com/EWS/Exchange.asmx";
	Map<String,Folder> mapKey=new HashMap<>();
	
	Map<String,Folder> mapKeyTree=new HashMap<>();
	
	private String accessToken;
	private String refreshToken;
	EWSAccesToken ewsAccesToken;
	
   public  ExchangeService loginEWS(String username,String password) {
       ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
       try {

           service.setCredentials(new WebCredentials(username, password));
          // service.autodiscoverUrl(username);
         //   service.autodiscoverUrl(username,new RedirectionUrlCallback());
           service.setUrl(new java.net.URI(HOST));    
           //System.out.println(service.getUrl());
           service.setTraceEnabled(Boolean.TRUE);
           service.validate();
           service.setTimeout(300000);
       }
       catch (Exception e){
           e.printStackTrace();
       }
       return service;
   }

   
   public  ExchangeService loginEWS(String username) {
       
       try {
    	   
    	    ewsAccesToken= new EWSAccesToken();
    	    accessToken= ewsAccesToken.getAccessToken();
    	    refreshToken= ewsAccesToken.getRefreshToken();
    	    service = new ExchangeService(ExchangeVersion.Exchange2013);
    	    service.getHttpHeaders().put("Authorization", "Bearer" + accessToken);
            service.getHttpHeaders().put("X-AnchorMailbox", username);
            service.setUrl(new java.net.URI(HOST));
           
             
          ImpersonatedUserId impersonatedUserId= service.getImpersonatedUserId();
      
       
          // service.setImpersonatedUserId(new ImpersonatedUserId(ConnectingIdType.PrincipalName, username));
          // service.setImpersonatedUserId(new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "dispatch@navinixllc.com"));
           service.setTraceEnabled(Boolean.TRUE);
           service.validate(); 
       }
       catch (Exception e){
           e.printStackTrace();
       }
       return service;
   }
  public  ExchangeService loginRefreshTokenEWS(String username,String refreshToken) {
       
       try {

    	    accessToken= ewsAccesToken.getRefreshAccessToken(refreshToken);
    	    refreshToken= ewsAccesToken.getRefreshToken();
    	    service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    	    service.getHttpHeaders().put("Authorization", "Bearer" + accessToken);
            service.getHttpHeaders().put("X-AnchorMailbox", username);
            service.setUrl(new java.net.URI(HOST));
          // System.out.println(service.getUrl());
           service.setTraceEnabled(Boolean.TRUE);
           service.validate();
          // service.setTimeout(300000); 
           
       }
       catch (Exception e){
           e.printStackTrace();
       }
       return service;
   }
   
   
   public String getRefreshToken()
	{
		return refreshToken;
	}
    
 public  Map<String,Folder>  getFolder(ExchangeService service,DefaultTableModel model,boolean deepFolderTraversal) throws Exception
 {
		FolderView folderView = new FolderView(Integer.MAX_VALUE);
//		if (deepFolderTraversal) {
//			folderView.setTraversal(FolderTraversal.Deep);
//		}

		FindFoldersResults findFolderResults = null;
		if (EmailWizardApplication.r_mailbox.isSelected()) {
			findFolderResults = service.findFolders(WellKnownFolderName.MsgFolderRoot, folderView);
		}
		if (EmailWizardApplication.r_public.isSelected()) {
			findFolderResults = service.findFolders(WellKnownFolderName.PublicFoldersRoot, folderView);
		}
		if (EmailWizardApplication.r_archive.isSelected()) {
			findFolderResults = service.findFolders(WellKnownFolderName.ArchiveMsgFolderRoot, folderView);
		}

		int count = 0;
		for (Folder folder : findFolderResults) {

			byte[] germanBytes = folder.getDisplayName().getBytes();

			String asciiEncodedString = new String(germanBytes, StandardCharsets.US_ASCII);
			mapKey.put(asciiEncodedString, folder);
			model.addRow(new Object[] { count, asciiEncodedString, folder.getTotalCount(), true });
			count++;
		}

		return mapKey;

 }
 public  Map<String,Folder>  getFolderTree(ExchangeService service,DefaultMutableTreeNode modelTree,DefaultTableModel model,boolean deepFolderTraversal) throws Exception
 {
		FolderView folderView = new FolderView(Integer.MAX_VALUE);
//		if (deepFolderTraversal) {
//			folderView.setTraversal(FolderTraversal.Deep);
//		}

		FindFoldersResults findFolderResults = null;
		if (EmailWizardApplication.r_mailbox.isSelected()) {
			findFolderResults = service.findFolders(WellKnownFolderName.MsgFolderRoot, folderView);
			
		}
		if (EmailWizardApplication.r_public.isSelected()) {
			findFolderResults = service.findFolders(WellKnownFolderName.PublicFoldersRoot, folderView);
		}
		if (EmailWizardApplication.r_archive.isSelected()) {
			findFolderResults = service.findFolders(WellKnownFolderName.ArchiveMsgFolderRoot, folderView);
		}
		
		int count = 0;
		for (Folder folder : findFolderResults) {

			byte[] germanBytes = folder.getDisplayName().getBytes();

			String asciiEncodedString = new String(germanBytes, StandardCharsets.US_ASCII);
			mapKey.put(asciiEncodedString, folder);
			model.addRow(new Object[] { count, asciiEncodedString, folder.getTotalCount(), true });
			count++;
		}
		
		getSubFolders(service,findFolderResults,modelTree);	
		
		return mapKeyTree;

 }
 

 public Map<String,Folder> setMapKey(Map<String,Folder> mapKey)
 {
	 return this.mapKey=mapKey;
 }
 public Map<String,Folder> getMapKey()
 {
	 return this.mapKey;
 }
 
 public Map<String,Folder> getMapKeyTree()
 {
	 return this.mapKeyTree;
 }
 
 public void getSubFolders(ExchangeService service,FindFoldersResults findFolderResults,DefaultMutableTreeNode node) throws Exception
 {
		for (Folder folder : findFolderResults) {
			

			if (folder.getChildFolderCount() > 0) {
				
				DefaultMutableTreeNode subNode=new DefaultMutableTreeNode(folder.getDisplayName());
				node.add(subNode);
				mapKeyTree.put(folder.getDisplayName(), folder);
				
				
				findFolderResults = service.findFolders(folder.getId(), new FolderView(Integer.MAX_VALUE));
				getSubFolders(service,findFolderResults,subNode);
				
			} else {
				System.out.println(folder.getDisplayName());			
				mapKeyTree.put(folder.getDisplayName(), folder);
				DefaultMutableTreeNode subNode=new DefaultMutableTreeNode(folder.getDisplayName());
				node.add(subNode);
				
			}

		}	 
 }
	
 public static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
     public boolean autodiscoverRedirectionUrlValidationCallback(String redirectionUrl) {
       return redirectionUrl.toLowerCase().startsWith("https://");
     }
 }


}