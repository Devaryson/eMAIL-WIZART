package com.util;

import java.io.ByteArrayInputStream;
import java.util.Date;

import com.aspose.email.MailMessage;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiProperty;
import com.aspose.email.MapiPropertyTag;
import com.aspose.email.MapiTask;
import com.aspose.email.TIPMethod;
import com.aspose.email.system.DateTime;

import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Contact;
import microsoft.exchange.webservices.data.core.service.item.Task;

public interface MapiMessageUtils {

	public static MapiTask setTaskMapiProperty(MapiTask mapiTask, Task task) throws ServiceLocalException {
		DateTime PR_CREATION_TIME = new DateTime();
		PR_CREATION_TIME = PR_CREATION_TIME.fromJava(task.getDateTimeCreated());
		MapiProperty propertyCreationTime = new MapiProperty(MapiPropertyTag.PR_CREATION_TIME,
				convertDateTime(PR_CREATION_TIME));
		mapiTask.setProperty(propertyCreationTime);

		DateTime PR_RECEIPT_TIME = new DateTime();
		PR_RECEIPT_TIME = PR_RECEIPT_TIME.fromJava(task.getDateTimeReceived());
		MapiProperty propertyReceipt = new MapiProperty(MapiPropertyTag.PR_RECEIPT_TIME,
				convertDateTime(PR_RECEIPT_TIME));
		mapiTask.setProperty(propertyReceipt);

		DateTime PR_LAST_MODIFICATION_TIME = new DateTime();
		PR_LAST_MODIFICATION_TIME = PR_LAST_MODIFICATION_TIME.fromJava(task.getLastModifiedTime());
		MapiProperty modificationTime = new MapiProperty(MapiPropertyTag.PR_LAST_MODIFICATION_TIME,
				convertDateTime(PR_LAST_MODIFICATION_TIME));
		mapiTask.setProperty(modificationTime);

		MapiProperty PR_MESSAGE_CLASS = new MapiProperty(MapiPropertyTag.PR_MESSAGE_CLASS,
				task.getItemClass().getBytes());
		mapiTask.setProperty(PR_MESSAGE_CLASS);
		return mapiTask;

	}

	public static MapiContact setContactMapiProperty(MapiContact mapiContact, Contact contact)
			throws ServiceLocalException {
		DateTime PR_CREATION_TIME = new DateTime();
		PR_CREATION_TIME = PR_CREATION_TIME.fromJava(contact.getDateTimeCreated());
		MapiProperty propertyCreationTime = new MapiProperty(MapiPropertyTag.PR_CREATION_TIME,
				convertDateTime(PR_CREATION_TIME));

		mapiContact.setProperty(propertyCreationTime);

		DateTime PR_RECEIPT_TIME = new DateTime();
		PR_RECEIPT_TIME = PR_RECEIPT_TIME.fromJava(contact.getDateTimeReceived());
		MapiProperty propertyReceipt = new MapiProperty(MapiPropertyTag.PR_RECEIPT_TIME,
				convertDateTime(PR_RECEIPT_TIME));

		mapiContact.setProperty(propertyReceipt);

		DateTime PR_LAST_MODIFICATION_TIME = new DateTime();
		PR_LAST_MODIFICATION_TIME = PR_LAST_MODIFICATION_TIME.fromJava(contact.getLastModifiedTime());
		MapiProperty modificationTime = new MapiProperty(MapiPropertyTag.PR_LAST_MODIFICATION_TIME,
				convertDateTime(PR_LAST_MODIFICATION_TIME));

		mapiContact.setProperty(modificationTime);

		try {
			MapiProperty subject = new MapiProperty(MapiPropertyTag.PR_SUBJECT, contact.getDisplayName().getBytes());
			mapiContact.setProperty(subject);
		} catch (Exception ex) {

			// logger.error("An exception occurred!", ex);
		}

		MapiProperty PR_MESSAGE_CLASS = new MapiProperty(MapiPropertyTag.PR_MESSAGE_CLASS,
				contact.getItemClass().getBytes());
		mapiContact.setProperty(PR_MESSAGE_CLASS);

		return mapiContact;

	}

	public static MapiTask convertToMapiTask(ByteArrayInputStream bis, Task task) throws ServiceLocalException {

		com.aspose.email.Task asposeTask = new com.aspose.email.Task();
		asposeTask.setMethod(TIPMethod.Request);
		asposeTask.setSubject(task.getSubject());
		asposeTask.setBody(task.getBody().toString());
		if (task.getStartDate() != null & task.getStartDate() != null) {
			asposeTask.setStartDate(new Date(task.getStartDate().getTime()));
			asposeTask.setDueDate(new Date(task.getDueDate().getTime()));
		}
		MailMessage mailMessage = new MailMessage().load(bis);
		mailMessage.addAlternateView(asposeTask.request());
		

		MapiConversionOptions m = MapiConversionOptions.getUnicodeFormat();
		m.setPreserveEmbeddedMessageFormat(true);
		m.setForcedRtfBodyForAppointment(true);
		m.setPreserveOriginalAddresses(true);
		m.setPreserveOriginalDates(true);

		MapiMessage mapiMsg = MapiMessage.fromMailMessage(mailMessage, m);
		mailMessage.close();
		MapiTask mapiTask = (MapiTask) mapiMsg.toMapiMessageItem();
		mapiMsg.close();

		return mapiTask;
	}

	public static byte[] convertDateTime(DateTime t) {
		long filetime = t.toFileTime();
		byte[] d = new byte[8];
		d[0] = (byte) (filetime & 0xFF);
		d[1] = (byte) ((filetime & 0xFF00) >> 8);
		d[2] = (byte) ((filetime & 0xFF0000) >> 16);
		d[3] = (byte) ((filetime & 0xFF000000) >> 24);
		d[4] = (byte) ((filetime & 0xFF00000000l) >> 32);
		d[5] = (byte) ((filetime & 0xFF0000000000l) >> 40);
		d[6] = (byte) ((filetime & 0xFF000000000000l) >> 48);
		d[7] = (byte) ((filetime & 0xFF00000000000000l) >> 56);
		return d;
	}

	public static MapiCalendar setCalendarMapiProperty(MapiCalendar mapiCalendar, Appointment appoinment)
			throws ServiceLocalException {
		DateTime PR_CREATION_TIME = new DateTime();
		PR_CREATION_TIME = PR_CREATION_TIME.fromJava(appoinment.getDateTimeCreated());
		MapiProperty propertyCreationTime = new MapiProperty(MapiPropertyTag.PR_CREATION_TIME,
				convertDateTime(PR_CREATION_TIME));
		mapiCalendar.setProperty(propertyCreationTime);

		DateTime PR_RECEIPT_TIME = new DateTime();
		PR_RECEIPT_TIME = PR_RECEIPT_TIME.fromJava(appoinment.getDateTimeReceived());
		MapiProperty propertyReceipt = new MapiProperty(MapiPropertyTag.PR_RECEIPT_TIME,
				convertDateTime(PR_RECEIPT_TIME));
		mapiCalendar.setProperty(propertyReceipt);

		DateTime PR_LAST_MODIFICATION_TIME = new DateTime();
		PR_LAST_MODIFICATION_TIME = PR_LAST_MODIFICATION_TIME.fromJava(appoinment.getLastModifiedTime());
		MapiProperty modificationTime = new MapiProperty(MapiPropertyTag.PR_LAST_MODIFICATION_TIME,
				convertDateTime(PR_LAST_MODIFICATION_TIME));
		mapiCalendar.setProperty(modificationTime);

		MapiProperty PR_MESSAGE_CLASS = new MapiProperty(MapiPropertyTag.PR_MESSAGE_CLASS,
				appoinment.getItemClass().getBytes());
		mapiCalendar.setProperty(PR_MESSAGE_CLASS);
		
		

		return mapiCalendar;

	}

}
