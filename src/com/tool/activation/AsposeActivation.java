package com.tool.activation;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import com.main.EmailWizardApplication;

public class AsposeActivation {

	public void doAsposeLicActivation()
	{
		try {
			InputStream is = EmailWizardApplication.class.getResourceAsStream("/data.txt");
			byte[] buff = new byte[is.available()];
			is.read(buff);
			ency td = new ency();
			String decrpt = td.decrypt(buff);
			OutputStream outStream = new FileOutputStream(
					System.getProperty("java.io.tmpdir") + File.separator + "Email.txt");
			outStream.write(decrpt.getBytes());
			is.close();
			outStream.flush();
			outStream.close();
			File file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Email.txt");
			if (file.exists()) {
				com.aspose.email.License lic_email = new com.aspose.email.License();
				com.aspose.words.License lic_word = new com.aspose.words.License();
				com.aspose.pdf.License lic_pdf = new com.aspose.pdf.License();
				com.aspose.cells.License lic_cell = new com.aspose.cells.License();
				lic_email.setLicense(file.getPath());
				lic_word.setLicense(file.getPath());
				lic_pdf.setLicense(file.getPath());
				lic_cell.setLicense(file.getPath());
				file.delete();
			}
			file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Email.txt");
			if (file.exists()) {
				file.delete();
			}
		} catch (Exception e1) {
			e1.printStackTrace();

		}
	}
}
