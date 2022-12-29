package com.util;

import java.io.File;

import javax.swing.table.DefaultTableModel;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.main.EmailWizardApplication;

public interface CSVUtils {
	
	public static Workbook createSampleCSVStructure() {
		Workbook workbook = new Workbook();
		Cell a1 = workbook.getWorksheets().get(0).getCells().get("A1");
		a1.setValue("Email ID");
		Cell b1 = workbook.getWorksheets().get(0).getCells().get("B1");
		b1.setValue("Password");
		Cell c1 = workbook.getWorksheets().get(0).getCells().get("C1");
		c1.setValue("Imap Host");
		Cell d1 = workbook.getWorksheets().get(0).getCells().get("D1");
		d1.setValue("Port No");
		return workbook;
		
	}

	public static Workbook createCSVStructure(int cellNo) {
		Workbook workbook = new Workbook();
		Cell a1 = workbook.getWorksheets().get(0).getCells().get("A" + cellNo);
		a1.setValue("Subject");
		Cell b1 = workbook.getWorksheets().get(0).getCells().get("B" + cellNo);
		b1.setValue("From");
		Cell c1 = workbook.getWorksheets().get(0).getCells().get("C" + cellNo);
		c1.setValue("Body");
		Cell d1 = workbook.getWorksheets().get(0).getCells().get("D" + cellNo);
		d1.setValue("To");
		Cell e1 = workbook.getWorksheets().get(0).getCells().get("E" + cellNo);
		e1.setValue("Date");
		Cell f1 = workbook.getWorksheets().get(0).getCells().get("F" + cellNo);
		f1.setValue("Bcc");
		Cell g1 = workbook.getWorksheets().get(0).getCells().get("G" + cellNo);
		g1.setValue("Cc");
		Cell h1 = workbook.getWorksheets().get(0).getCells().get("H" + cellNo);
		h1.setValue("StartDate");
		Cell i1 = workbook.getWorksheets().get(0).getCells().get("I" + cellNo);
		i1.setValue("EndDate");
		Cell j1 = workbook.getWorksheets().get(0).getCells().get("J" + cellNo);
		j1.setValue("Location");
		return workbook;
	}

	public static Workbook createCSVStructureContact(int cellNo) {
		Workbook workbook = new Workbook();
		Cell a1 = workbook.getWorksheets().get(0).getCells().get("A" + cellNo);
		a1.setValue("Title");
		Cell b1 = workbook.getWorksheets().get(0).getCells().get("B" + cellNo);
		b1.setValue("First Name");
		Cell c1 = workbook.getWorksheets().get(0).getCells().get("C" + cellNo);
		c1.setValue("Middle Name");
		Cell d1 = workbook.getWorksheets().get(0).getCells().get("D" + cellNo);
		d1.setValue("Last Name");

		Cell e1 = workbook.getWorksheets().get(0).getCells().get("E" + cellNo);
		e1.setValue("Suffix");
		Cell f1 = workbook.getWorksheets().get(0).getCells().get("F" + cellNo);
		f1.setValue("Company");
		Cell g1 = workbook.getWorksheets().get(0).getCells().get("G" + cellNo);
		g1.setValue("Department");
		Cell h1 = workbook.getWorksheets().get(0).getCells().get("H" + cellNo);
		h1.setValue("Job Title");

		Cell i1 = workbook.getWorksheets().get(0).getCells().get("I" + cellNo);
		i1.setValue("Business Street");
		Cell j1 = workbook.getWorksheets().get(0).getCells().get("J" + cellNo);
		j1.setValue("Business Street 2");
		Cell k1 = workbook.getWorksheets().get(0).getCells().get("k" + cellNo);
		k1.setValue("Business Street 3");
		Cell l1 = workbook.getWorksheets().get(0).getCells().get("L" + cellNo);
		l1.setValue("Business City");
		Cell m1 = workbook.getWorksheets().get(0).getCells().get("M" + cellNo);
		m1.setValue("Business State");
		Cell n1 = workbook.getWorksheets().get(0).getCells().get("N" + cellNo);
		n1.setValue("Business Postal Code");
		Cell o1 = workbook.getWorksheets().get(0).getCells().get("O" + cellNo);
		o1.setValue("Business Country/Region");

		Cell p1 = workbook.getWorksheets().get(0).getCells().get("P" + cellNo);
		p1.setValue("Home Street");
		Cell q1 = workbook.getWorksheets().get(0).getCells().get("Q" + cellNo);
		q1.setValue("Home Street 2");
		Cell r1 = workbook.getWorksheets().get(0).getCells().get("R" + cellNo);
		r1.setValue("Home Street 3");
		Cell s1 = workbook.getWorksheets().get(0).getCells().get("S" + cellNo);
		s1.setValue("Home City");
		Cell t1 = workbook.getWorksheets().get(0).getCells().get("T" + cellNo);
		t1.setValue("Home State");
		Cell u1 = workbook.getWorksheets().get(0).getCells().get("U" + cellNo);
		u1.setValue("Home Postal Code");
		Cell v1 = workbook.getWorksheets().get(0).getCells().get("V" + cellNo);
		v1.setValue("Home Country/Region");

		Cell w1 = workbook.getWorksheets().get(0).getCells().get("W" + cellNo);
		w1.setValue("Other Street");
		Cell x1 = workbook.getWorksheets().get(0).getCells().get("X" + cellNo);
		x1.setValue("Other City");
		Cell y1 = workbook.getWorksheets().get(0).getCells().get("Y" + cellNo);
		y1.setValue("Other State");
		Cell z1 = workbook.getWorksheets().get(0).getCells().get("Z" + cellNo);
		z1.setValue("Other Country/Region");

		Cell aa = workbook.getWorksheets().get(0).getCells().get("AA" + cellNo);
		aa.setValue("Assistant's Phone");
		Cell ab = workbook.getWorksheets().get(0).getCells().get("AB" + cellNo);
		ab.setValue("Home Phone");
		Cell ac = workbook.getWorksheets().get(0).getCells().get("AC" + cellNo);
		ac.setValue("ISDN");
		Cell ad = workbook.getWorksheets().get(0).getCells().get("AD" + cellNo);
		ad.setValue("Mobile Phone");

		Cell ae = workbook.getWorksheets().get(0).getCells().get("AE" + cellNo);
		ae.setValue("Anniversary");
		Cell af = workbook.getWorksheets().get(0).getCells().get("AF" + cellNo);
		af.setValue("Birthday");

		Cell ag = workbook.getWorksheets().get(0).getCells().get("AG" + cellNo);
		ag.setValue("E-mail Address");
		Cell ah = workbook.getWorksheets().get(0).getCells().get("AH" + cellNo);
		ah.setValue("Body");

		Cell ai = workbook.getWorksheets().get(0).getCells().get("AI" + cellNo);
		ai.setValue("Gender");

		Cell aj = workbook.getWorksheets().get(0).getCells().get("AJ" + cellNo);
		aj.setValue("Notes");

		Cell ak = workbook.getWorksheets().get(0).getCells().get("AK" + cellNo);
		ak.setValue("Home Email Address");

		Cell al = workbook.getWorksheets().get(0).getCells().get("AL" + cellNo);
		al.setValue("Work Email Address");

		Cell am = workbook.getWorksheets().get(0).getCells().get("AM" + cellNo);
		am.setValue(" Email Address 1");

		Cell an = workbook.getWorksheets().get(0).getCells().get("AN" + cellNo);
		an.setValue(" Email Address 2");

		Cell ao = workbook.getWorksheets().get(0).getCells().get("AO" + cellNo);
		ao.setValue(" Email Address 3");

		return workbook;
	}

	public static Workbook saveCSVContact(Workbook workbook, int cellNo, MapiContact mapiContact) {
		workbook.getWorksheets().get(0).getCells().get("A" + cellNo).putValue(mapiContact.getSubject());
		workbook.getWorksheets().get(0).getCells().get("B" + cellNo).putValue(mapiContact.getNameInfo().getGivenName());
		workbook.getWorksheets().get(0).getCells().get("C" + cellNo)
				.putValue(mapiContact.getNameInfo().getMiddleName());
		workbook.getWorksheets().get(0).getCells().get("D" + cellNo).putValue(mapiContact.getNameInfo().getSurname());
		workbook.getWorksheets().get(0).getCells().get("E" + cellNo)
				.putValue(mapiContact.getNameInfo().getDisplayName());
		workbook.getWorksheets().get(0).getCells().get("F" + cellNo)
				.putValue(mapiContact.getProfessionalInfo().getCompanyName());

		workbook.getWorksheets().get(0).getCells().get("G" + cellNo)
				.putValue(mapiContact.getProfessionalInfo().getDepartmentName());

		workbook.getWorksheets().get(0).getCells().get("H" + cellNo)
				.putValue(mapiContact.getProfessionalInfo().getTitle());

		workbook.getWorksheets().get(0).getCells().get("I" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getWorkAddress().getStreet());

		workbook.getWorksheets().get(0).getCells().get("J" + cellNo).putValue("");

		workbook.getWorksheets().get(0).getCells().get("K" + cellNo).putValue("");

		workbook.getWorksheets().get(0).getCells().get("L" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getWorkAddress().getCity());

		workbook.getWorksheets().get(0).getCells().get("M" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getWorkAddress().getStateOrProvince());

		workbook.getWorksheets().get(0).getCells().get("N" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getWorkAddress().getPostalCode());

		workbook.getWorksheets().get(0).getCells().get("O" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getWorkAddress().getCountry());

		workbook.getWorksheets().get(0).getCells().get("P" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getHomeAddress().getStreet());

		workbook.getWorksheets().get(0).getCells().get("Q" + cellNo).putValue("");
		workbook.getWorksheets().get(0).getCells().get("R" + cellNo).putValue("");

		workbook.getWorksheets().get(0).getCells().get("S" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getHomeAddress().getCity());

		workbook.getWorksheets().get(0).getCells().get("T" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getHomeAddress().getStateOrProvince());

		workbook.getWorksheets().get(0).getCells().get("U" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getHomeAddress().getPostalCode());

		workbook.getWorksheets().get(0).getCells().get("V" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getHomeAddress().getCountry());

		workbook.getWorksheets().get(0).getCells().get("W" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getOtherAddress().getStreet());

		workbook.getWorksheets().get(0).getCells().get("X" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getOtherAddress().getCity());

		workbook.getWorksheets().get(0).getCells().get("Y" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getOtherAddress().getStateOrProvince());
		workbook.getWorksheets().get(0).getCells().get("Z" + cellNo)
				.putValue(mapiContact.getPhysicalAddresses().getOtherAddress().getCountry());

		workbook.getWorksheets().get(0).getCells().get("AA" + cellNo)
				.putValue(mapiContact.getTelephones().getAssistantTelephoneNumber());

		workbook.getWorksheets().get(0).getCells().get("AB" + cellNo)
				.putValue(mapiContact.getTelephones().getHomeTelephoneNumber());

		workbook.getWorksheets().get(0).getCells().get("AC" + cellNo)
				.putValue(mapiContact.getTelephones().getIsdnNumber());

		workbook.getWorksheets().get(0).getCells().get("AD" + cellNo)
				.putValue(mapiContact.getTelephones().getMobileTelephoneNumber());

		workbook.getWorksheets().get(0).getCells().get("AE" + cellNo)
				.putValue(mapiContact.getEvents().getWeddingAnniversary());

		workbook.getWorksheets().get(0).getCells().get("AF" + cellNo).putValue(mapiContact.getEvents().getBirthday());

		workbook.getWorksheets().get(0).getCells().get("AG" + cellNo)
				.putValue(mapiContact.getElectronicAddresses().getDefaultEmailAddress());

		workbook.getWorksheets().get(0).getCells().get("AH" + cellNo).putValue(mapiContact.getBody());

		workbook.getWorksheets().get(0).getCells().get("AI" + cellNo)
				.putValue(mapiContact.getPersonalInfo().getGender());

		workbook.getWorksheets().get(0).getCells().get("AJ" + cellNo)
				.putValue(mapiContact.getPersonalInfo().getNotes());

		workbook.getWorksheets().get(0).getCells().get("AK" + cellNo)
				.putValue(mapiContact.getElectronicAddresses().getHomeFax());

		workbook.getWorksheets().get(0).getCells().get("AL" + cellNo)
				.putValue(mapiContact.getElectronicAddresses().getPrimaryFax());

		workbook.getWorksheets().get(0).getCells().get("AM" + cellNo)
				.putValue(mapiContact.getElectronicAddresses().getEmail1());

		workbook.getWorksheets().get(0).getCells().get("AN" + cellNo)
				.putValue(mapiContact.getElectronicAddresses().getEmail2());

		workbook.getWorksheets().get(0).getCells().get("AO" + cellNo)
				.putValue(mapiContact.getElectronicAddresses().getEmail3());
		return workbook;

	}

	public static Workbook saveCSVEmailandTask(Workbook workbook, int cellNo, MailMessage msg) {
		String body = msg.getBody();
		if (body.length() > 32000) {
			body = body.substring(0, 31000);
		}
		workbook.getWorksheets().get(0).getCells().get("A" + cellNo).putValue(msg.getSubject());
		workbook.getWorksheets().get(0).getCells().get("B" + cellNo).putValue(msg.getFrom());
		workbook.getWorksheets().get(0).getCells().get("C" + cellNo).putValue(body);
		workbook.getWorksheets().get(0).getCells().get("D" + cellNo).putValue(msg.getTo());
		workbook.getWorksheets().get(0).getCells().get("E" + cellNo).putValue(msg.getDate());
		workbook.getWorksheets().get(0).getCells().get("F" + cellNo).putValue(msg.getBcc());
		workbook.getWorksheets().get(0).getCells().get("G" + cellNo).putValue(msg.getCC());
		return workbook;
	}

	public static Workbook saveCSVAppointment(Workbook workbook, int cellNo, MapiCalendar mapiCalendar) {
		String body = mapiCalendar.getBody();
		if (body.length() > 32000) {
			body = body.substring(0, 31000);
		}
		workbook.getWorksheets().get(0).getCells().get("A" + cellNo).putValue(mapiCalendar.getSubject());
		workbook.getWorksheets().get(0).getCells().get("B" + cellNo).putValue(mapiCalendar.getOrganizer());
		workbook.getWorksheets().get(0).getCells().get("C" + cellNo).putValue(body);
		workbook.getWorksheets().get(0).getCells().get("D" + cellNo).putValue(mapiCalendar.getAttendees());
		workbook.getWorksheets().get(0).getCells().get("H" + cellNo).putValue(mapiCalendar.getStartDate());
		workbook.getWorksheets().get(0).getCells().get("I" + cellNo).putValue(mapiCalendar.getEndDate());
		workbook.getWorksheets().get(0).getCells().get("J" + cellNo).putValue(mapiCalendar.getLocation());
		return workbook;
	}
	public static Workbook saveCSV(Workbook workbook, File finalDestination) {
		try {
			workbook.save(finalDestination.getAbsolutePath() + File.separator + finalDestination.getName() + ".csv",com.aspose.cells.SaveFormat.CSV);					
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return workbook;
	}
	public static void readCSV(DefaultTableModel loginTableModel, String csvDestination) {
		try {
		    Workbook workbook = new Workbook(csvDestination);
			Worksheet sheet = workbook.getWorksheets().get(0);						
			int csvRowCount = sheet.getCells().getRows().getCount();
			int tableRowCount=0;
			while (loginTableModel.getRowCount() > 0) {loginTableModel.removeRow(0);}
			for (int i = 2; i <=csvRowCount; i++) {
				String cellEmailID =  sheet.getCells().get("A"+i).getValue().toString().trim();					
				String cellPassword = sheet.getCells().get("B"+i).getValue().toString().trim();					
				String cellHostName=  sheet.getCells().get("C"+i).getValue().toString().trim();			
				String cellPortNo =   sheet.getCells().get("D"+i).getValue().toString().trim();						 
				loginTableModel.addRow(new Object[] { tableRowCount,cellEmailID, cellPassword,cellHostName, cellPortNo, 0 });
				tableRowCount++;
			}				
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
