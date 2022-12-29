package com.aspose.word;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.file.CopyOption;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

import org.apache.commons.io.FileSystemUtils;
import org.apache.commons.io.FileUtils;

import com.aspose.cells.Encoding;
import com.aspose.email.Attachment;
import com.aspose.email.AttachmentCollection;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MapiMessage;
import com.aspose.email.SaveOptions;
import com.aspose.pdf.facades.PdfContentEditor;
import com.aspose.pdf.internal.imaging.CharacterSet;
import com.aspose.words.Document;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;
import com.main.EmailWizardApplication;

public class DownloadWordFormat {
	

	
	public void saveWordFormat(MailMessage msg, String destinationPath) throws Exception {
		
		 if(EmailWizardApplication.chckbxSaveSeperateAttachments.isSelected())
		{						 			
			destinationPath= seperateAttachments(msg,destinationPath);		
		}
		
		if (EmailWizardApplication.r_Eml.isSelected()) {
			msg.save(destinationPath + ".eml", SaveOptions.getDefaultEml());
			  
		} else if (EmailWizardApplication.r_msg.isSelected()) {

			MapiMessage mapi = MapiMessage.fromMailMessage(msg);
			mapi.save(destinationPath + ".msg", SaveOptions.getDefaultMsgUnicode());
	  
		} else if (EmailWizardApplication.r_emlx.isSelected()) {

			msg.save(destinationPath + ".emlx", SaveOptions.createSaveOptions(MailMessageSaveType.getEmlxFormat()));
			  

		} else if (EmailWizardApplication.r_html.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".html", SaveFormat.HTML);
			  
		} else if (EmailWizardApplication.r_rtf.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".rtf", SaveFormat.RTF);

		} else if (EmailWizardApplication.r_xps.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".xps", SaveFormat.XPS);

		} else if (EmailWizardApplication.r_emf.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".emf", SaveFormat.EMF);

		} else if (EmailWizardApplication.r_docx.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".docx", SaveFormat.DOCX);

		} else if (EmailWizardApplication.r_jpeg.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".jpeg", SaveFormat.JPEG);

		} else if (EmailWizardApplication.r_docm.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".docm", SaveFormat.DOCM);

		} else if (EmailWizardApplication.r_text.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".text", SaveFormat.TEXT);

		} else if (EmailWizardApplication.r_tiff.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".tiff", SaveFormat.TIFF);

		} else if (EmailWizardApplication.r_png.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".png", SaveFormat.PNG);

		}
		else if (EmailWizardApplication.r_svg.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".svg", SaveFormat.SVG);
		}

		else if (EmailWizardApplication.r_epub.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".epub", SaveFormat.EPUB);

		} else if (EmailWizardApplication.r_dotm.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".dotm", SaveFormat.DOTM);

		} else if (EmailWizardApplication.r_ott.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".ott", SaveFormat.OTT);

		} else if (EmailWizardApplication.r_gif.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".gif", SaveFormat.GIF);

		} else if (EmailWizardApplication.r_bmp.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".bmp", SaveFormat.BMP);

		}

		else if (EmailWizardApplication.r_wordml.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".wordml", SaveFormat.WORD_ML);

		}

		else if (EmailWizardApplication.r_odt.isSelected()) {

			Document document = convertMSGToDocument(msg);
			document.save(destinationPath + ".odt", SaveFormat.ODT);
		}
        else if (EmailWizardApplication.r_pdf.isSelected()) {
			
			Document document = convertMSGToDocument(msg);
			
			document.save(destinationPath + ".pdf", SaveFormat.PDF);
			AttachmentCollection attachmentCollection = msg.getAttachments();

			for (Attachment attachment : attachmentCollection) {
				if (attachment != null) {
					PdfContentEditor editor = new PdfContentEditor();
					editor.bindPdf(destinationPath + ".pdf");
					if (attachment.getName() != null && attachment.getName().length() > 0) {

						InputStream is = attachment.getContentStream();
						editor.addDocumentAttachment(is, attachment.getName(), "");
						editor.save(destinationPath + ".pdf");
						is.close();
						editor.close();

					}
				}

			}
			document.cleanup();
		}
		
	}
	public Document convertMSGToDocument(MailMessage msg) throws Exception {
		LoadOptions lo = new LoadOptions();
		lo.setLoadFormat(LoadFormat.MHTML);
		//lo.setEncoding(Charset.forName("ISO-8859-1"));
		lo.setEncoding(Charset.forName("UTF-8"));
		lo.setPreserveIncludePictureField(true);

		long date=msg.getDate().getTime()+19800000;
		
	      msg.setDate(new Date(date));
		
		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
		msg.save(emlStream, SaveOptions.getDefaultMhtml());

		Document document = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
	
		emlStream.close();		  
		return document;
	}
	public String seperateAttachments(MailMessage msg,String destinationPath) throws Exception
	{
		MailMessage msgAttachment=msg.deepClone();
		
		File subjectNameFolder = new File(destinationPath);
		subjectNameFolder.mkdirs();
		
		File attachmentFolder = new File(destinationPath + File.separator + "Attachments");
		attachmentFolder.mkdirs();
		
		destinationPath = attachmentFolder.getAbsolutePath();

		saveAttachments(msgAttachment, destinationPath);

		destinationPath = subjectNameFolder.getAbsolutePath() + File.separator + subjectNameFolder.getName();
		msgAttachment.close();
		
		return destinationPath;
	}
	public void saveAttachments(MailMessage msg, String destinationPath)  throws Exception
	{
		  int attachmentCount=0;
			for (Attachment a : msg.getAttachments()) {
						
				    String attachmentName=namingConvention(a.getName());
				    File attachmnetTargetLocation = new File(destinationPath+File.separator+attachmentCount+"-"+attachmentName);
				   
					try(InputStream in =  a.getContentStream();) 
					{
						FileUtils.copyInputStreamToFile(in, attachmnetTargetLocation);
					
						
					}catch(FileAlreadyExistsException e) {
					    //destination file already exists
						System.out.println("File already exists exception...");
					} catch (IOException e) {
					    //something else went wrong
					    e.printStackTrace();
					}
				}
	                       
	} 
	

	
	String namingConvention(String subject) {

		String subjectName = subject;

		if (subjectName.length() > 40) {
			subjectName = subject.substring(0, 40);
		}

		return getRidOfIllegalFileNameCharacters(subjectName).trim();

	}
	
	static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
				.replace("//s", "").replace("\"", "");
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName;
	}

}
