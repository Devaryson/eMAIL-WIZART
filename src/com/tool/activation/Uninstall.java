package com.tool.activation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Base64;

import javax.swing.ImageIcon;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;

import com.main.EmailWizardApplication;
import com.tool.info.ToolDetails;

public class Uninstall {
	static File licFile;

	 String ToolUri;
	 String messageboxtitle;
	 String activationKeyno;
	 String orderid;
	 String type;
	 File licFileon;
	 File licFileonoff;
	 HttpURLConnection conn=null;
	
	
	public Uninstall(String ToolUri,String messageboxtitle,String activationKey,String orderId)
	{
		this.ToolUri=ToolUri;
		this.messageboxtitle=messageboxtitle;
		this.activationKeyno=activationKey;
		this.orderid=orderId;
		
		
        String fileKey = null;
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {
			licFileon = new File(System.getenv("APPDATA") + File.separator + ToolDetails.messageboxtitle + File.separator
					+ "licenseOnline.lic");	
			
			licFileonoff = new File(System.getenv("APPDATA") + File.separator + ToolDetails.messageboxtitle  + File.separator
					+ "license.lic");	
			
		} else {
			licFileon = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
					+ "Application Support" + File.separator + ToolDetails.messageboxtitle  + File.separator
					+ "licenseOnline.lic");
			licFileonoff = new File(System.getenv("APPDATA") + File.separator + ToolDetails.messageboxtitle  + File.separator
					+ "license.lic");	
		}
		if (licFileon.exists()) {
			try {
				
				FileReader fr=null;
				BufferedReader br =null;
				try
				{
				fr = new FileReader(licFileon);
				 br = new BufferedReader(new FileReader(licFileon));	
			
			
				int count=1;
				 while ((fileKey = br.readLine()) != null) {
					 System.out.println(fileKey);
					 if(count==1)
					 {
						 orderId =fileKey;							 
				byte[] actualByte = Base64.getDecoder().decode(orderId); 
	                       orderId = new String(actualByte); 
						 
					 }
					 else if(count==2)
					 {
						 activationKey =fileKey;
						 
							byte[] actualByte = Base64.getDecoder().decode(activationKey); 
							activationKey = new String(actualByte); 
						 
					 }
					 count++;
					    					
				  }
				}
				finally {
					
				
					br.close();
					fr.close();
				}
				
				orderid=orderId;
				activationKeyno=activationKey;
			
				
				if(!orderId.isEmpty()&&!activationKey.isEmpty())
				{
					//int version=Integer.valueOf(type);
				
						SwingUtilities.invokeLater(new Runnable() {

							@Override
							public void run() {										
								try
								{
								URL url = new URL(ToolUri+"order_no="+orderid+ "&activation_key="+activationKeyno +"&type=2");
								 conn = (HttpURLConnection) url.openConnection();					
				                 conn.setUseCaches(false);
								 conn.setDefaultUseCaches(false); 
								 conn.setRequestMethod("GET");				 
								 System.out.println(conn.getUseCaches()); 
								 
								String responsemessage = conn.getResponseMessage();
								int urlresponse=0;
								if (responsemessage.equalsIgnoreCase("ok")) {
									
									InputStream inputStream = conn.getInputStream();
									BufferedReader in = new BufferedReader(new InputStreamReader(inputStream));
									String inputLine;
									String response="";
									while ((inputLine = in.readLine()) != null) 
									{

										response=inputLine;
										System.out.println(inputLine);
									}
										 
									 urlresponse = Integer.valueOf(response);	
									 
									 if (urlresponse == 18) {											
											
										 System.out.println("Software Not Deactivate Due to some Error");
										 JOptionPane.showMessageDialog(null, "Software Not Deactivate Due to some Error", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
													new ImageIcon(Uninstall.class.getResource("/check.png")));
										 
										

										}
									 if (urlresponse == 17) {
										 
										 String path = "";

										
												if (System.getProperty("os.name").toLowerCase().contains("windows")) {

											
													path = System.getenv("APPDATA") + File.separator + ToolDetails.messageboxtitle ;

												} else {

												

													path = System.getProperty("user.home") + File.separator + "Library" + File.separator
															+ "Application Support" + File.separator + ToolDetails.messageboxtitle ;

												}

											


											

												File file = new File(path);
												recursiveDelete(file);
												System.out.println("Database deleted successfully...");
										 
										 
										 
										 
										 System.out.println("Deactivated Software Successfully");
										 JOptionPane.showMessageDialog(null, messageboxtitle+"  Deactivated Successfully", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
													new ImageIcon(Uninstall.class.getResource("/check.png")));
										 
									
										
									} else  if (urlresponse == 10) {
										
										System.out.println("Not Deactivated urlresponse ==10 ");
										
									}
										
								}
								else 
								{
									
									JOptionPane.showMessageDialog(null, "Unable To Fetch Response From the Server", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
											new ImageIcon(Uninstall.class.getResource("/information.png")));
									 
									
								}								
								}
							 catch (MalformedURLException e1) {
								
						
							    } 
								catch (IOException e1) {
							
									JOptionPane.showMessageDialog(null, "Not ready.Please check your Connection", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
											new ImageIcon(Uninstall.class.getResource("/information.png")));	
									
							}
									
							}
						});
						
						
			
						File file= new File(System.getenv("APPDATA") + File.separator + messageboxtitle);
						recursiveDelete(file);
					
				}
				else
				{
		
				}
				
			} catch (Exception ex) {
				
				ex.printStackTrace();


			}
		}
		else
		{
			
			JOptionPane.showMessageDialog(null, "This is Demo Version", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(Uninstall.class.getResource("/information.png")));	
			System.out.println(messageboxtitle+" Already uninstall");
		}
		
		
		
			}
	
	 public static void recursiveDelete(File file) {
	        //to end the recursive loop
	        if (!file.exists())
	            return;
	        
	        //if directory, go inside and call recursively
	        if (file.isDirectory()) {
	            for (File f : file.listFiles()) {
	                //call recursively
	                recursiveDelete(f);
	            }
	        }
	        //call delete to delete files and empty directory
	        file.delete();
	        System.out.println("Deleted file/folder: "+file.getAbsolutePath());
	    }
		}
