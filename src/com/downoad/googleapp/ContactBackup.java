package com.downoad.googleapp;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.security.GeneralSecurityException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import com.aspose.email.ContactSaveFormat;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactElectronicAddress;
import com.aspose.email.MapiContactElectronicAddressPropertySet;
import com.aspose.email.MapiContactEventPropertySet;
import com.aspose.email.MapiContactNamePropertySet;
import com.aspose.email.MapiContactPersonalInfoPropertySet;
import com.aspose.email.MapiContactPhysicalAddress;
import com.aspose.email.MapiContactPhysicalAddressPropertySet;
import com.aspose.email.MapiContactProfessionalPropertySet;
import com.aspose.email.MapiContactTelephonePropertySet;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.people.v1.PeopleService;
import com.google.api.services.people.v1.model.Address;
import com.google.api.services.people.v1.model.Birthday;
import com.google.api.services.people.v1.model.EmailAddress;
import com.google.api.services.people.v1.model.FieldMetadata;
import com.google.api.services.people.v1.model.Gender;
import com.google.api.services.people.v1.model.ImClient;
import com.google.api.services.people.v1.model.ListConnectionsResponse;
import com.google.api.services.people.v1.model.Name;
import com.google.api.services.people.v1.model.Occupation;
import com.google.api.services.people.v1.model.Organization;
import com.google.api.services.people.v1.model.Person;
import com.google.api.services.people.v1.model.PhoneNumber;
import com.google.api.services.people.v1.model.Skill;
import com.google.api.services.people.v1.model.Url;
import com.main.EmailWizardApplication;


public class ContactBackup {

	final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String APPLICATION_NAME = "Gmail Backup";
	static GoogleCredential gSuiteAppCredential;
	static PeopleService  contactService;
	
	@SuppressWarnings("deprecation")
	public  void downloadGsuiteContact(String serviceAccountId, String serviceAccountUser, String p12File) {
		
		try
		{
			final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			gSuiteAppCredential = new GoogleCredential.Builder().setTransport(HTTP_TRANSPORT).setJsonFactory(JSON_FACTORY)
					.setServiceAccountId(serviceAccountId)
					.setServiceAccountScopes(Collections.singleton("https://www.googleapis.com/auth/contacts.readonly"))
					.setServiceAccountUser(serviceAccountUser).setServiceAccountPrivateKeyFromP12File(new File(p12File))
					.build();
			 if (!gSuiteAppCredential.refreshToken()) {
			      throw new RuntimeException("Failed OAuth to refresh the token");
			    }
				contactService = new PeopleService.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(gSuiteAppCredential)).setApplicationName(APPLICATION_NAME).build();
						
				 downoadContacts();
		}
	      catch (GeneralSecurityException | IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	}
	public  void downloadGMailContact(Credential gMailAppCredential) {
		try
		{
			final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			contactService = new PeopleService.Builder(HTTP_TRANSPORT, JSON_FACTORY, setHttpTimeout(gMailAppCredential)).setApplicationName(APPLICATION_NAME).build();
			downoadContacts();
		}	
    catch (GeneralSecurityException | IOException e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
}
			
	}
	
	static boolean checkStopAndDemo(int i)
	{
		if (EmailWizardApplication.stop) {
			return true;
		}

		if (EmailWizardApplication.demo) {
			if (i == EmailWizardApplication.DEMO_LIMIT) {
				return true;
			}
		}
		return false;
		
	}
	@SuppressWarnings("deprecation")
	public  void downoadContacts() {

		try {

			EmailWizardApplication.progressBar_Downloading.setValue(0);
			EmailWizardApplication.progressBar_Downloading.setVisible(true);
			EmailWizardApplication.progressBar_Downloading.setMaximum(100);
			EmailWizardApplication.progressBar_Downloading.setStringPainted(true);
			EmailWizardApplication.lblDownloading.setVisible(true);
			
			
			ListConnectionsResponse response = contactService.people().connections().list("people/me")
				    //.setPersonFields(serviceAccountUser)
				
	                .setRequestMaskIncludeField("person.names,"
	                		+ "person.emailAddresses,"
	                		+ "person.phoneNumbers,"
	                		+ "person.addresses,"
	                		+ "person.birthdays,"
	                		+ "person.events,"
	                		+ "person.taglines,"
	                		+ "person.relations,"
	                	    + "person.relationship_statuses,"
	                		+ "person.residences,"
	                		+ "person.organizations,"
	                		+ "person.occupations,"
	                		+ "person.im_clients,"
	                		+ "person.nicknames,"
	                		+ "person.photos,"
	                		+ "person.residences,"
	                		+ "person.urls,"
	                		+ "person.genders")
            		
	                
		    .setPageSize(2000)
		    .execute();
		 // Display information about a person.
        List<Person> connections = response.getConnections();
        int count = 0;
        if (connections != null && connections.size() > 0) {
    for (Person person: connections){
 
                
		try {
			checkStopAndDemo(count);
			
			MapiContact contact = new MapiContact();


			// ------------------name------------------------------//
			String fullNameToDisplay = null;
			String NamePrefix = null;
			String givenNameToDisplay = null;
			String additionalNameToDisplay = null;
			String familyNameToDisplay = null;
			String hasNameSuffix = null;

			// ------------------Email-Address------------------------------//

			String EmailAddresses = null;
			String emailDisplayname = null;
			String userphoneNumber = null;

			// ------------------Phone Number------------------------------//
				List<PhoneNumber> phoneNumberList = person.getPhoneNumbers();
				if(phoneNumberList!=null)
				{
				MapiContactTelephonePropertySet Telephone = new MapiContactTelephonePropertySet();
				for (PhoneNumber phoneNumber : phoneNumberList) {

					if (phoneNumber!= null&&phoneNumber.getType()!=null&&phoneNumber.getValue()!=null) {

						
						String label=phoneNumber.getType();
						userphoneNumber = phoneNumber.getValue();						

						if (label.equalsIgnoreCase(("home"))) {
							Telephone.setHomeTelephoneNumber(userphoneNumber);
						} else if (label.equalsIgnoreCase(("work"))) {
							Telephone.setBusinessTelephoneNumber(userphoneNumber);
						} 
						else if (label.equalsIgnoreCase(("CONTACT"))) {
							Telephone.setMobileTelephoneNumber(userphoneNumber);
						} 						
						else if (label.equalsIgnoreCase(("other"))) {
							Telephone.setOtherTelephoneNumber(userphoneNumber);
						} else if (label.equalsIgnoreCase(("Work Fax"))) {
							Telephone.setBusiness2TelephoneNumber(userphoneNumber);
						} else if (label.equalsIgnoreCase(("Home Fax"))) {
							Telephone.setHome2TelephoneNumber(userphoneNumber);
						} else if (label.equalsIgnoreCase(("mobile"))) {
							Telephone.setMobileTelephoneNumber(userphoneNumber);
						} else if (label.equalsIgnoreCase(("main"))) {
							Telephone.setPrimaryTelephoneNumber(userphoneNumber);
						} else if (label.equalsIgnoreCase(("pager"))) {
							Telephone.setPagerTelephoneNumber(userphoneNumber);
						}
					}
					else if (phoneNumber!= null&&phoneNumber.getType()==null&&phoneNumber.getValue()!=null)
					{

						userphoneNumber = phoneNumber.getValue();
						Telephone.setMobileTelephoneNumber(userphoneNumber);
					}
					
				}
				contact.setTelephones(Telephone);
				}
				

				List<Name> nameList = person.getNames();
				if(nameList!=null)
				{
				for (Name name : nameList) {
					
					
				 if (name.getDisplayName()!=null) {
					fullNameToDisplay = name.getDisplayName();
				
				}

				 if (name.getHonorificPrefix()!=null) {

					NamePrefix = name.getHonorificPrefix();
			
				} 
				 if (name.getGivenName()!=null) {
					givenNameToDisplay = name.getGivenName();
					
				}
				

				 if (name.getMiddleName()!=null) {
					additionalNameToDisplay = name.getMiddleName();
					
				}

				 if (name.getFamilyName()!=null) {
					familyNameToDisplay = name.getFamilyName();
					
				} 
				 if (name.getHonorificSuffix()!=null) {
					hasNameSuffix = name.getHonorificSuffix();
				
				}
				}	
		}

			MapiContactNamePropertySet NamePropSet = new MapiContactNamePropertySet();
			NamePropSet.setDisplayName(fullNameToDisplay);
			NamePropSet.setMiddleName(additionalNameToDisplay);
			NamePropSet.setSurname(familyNameToDisplay);
			NamePropSet.setNickname(additionalNameToDisplay);
			NamePropSet.setInitials(givenNameToDisplay);
			NamePropSet.setDisplayNamePrefix(NamePrefix);
			NamePropSet.setGivenName(givenNameToDisplay);

			contact.setNameInfo(NamePropSet);

			//System.out.println("Email addresses:");

			MapiContactElectronicAddressPropertySet ElecAddrPropSet = new MapiContactElectronicAddressPropertySet();
			List<EmailAddress> emailList=person.getEmailAddresses();
			if(emailList!=null)
			{
			for (EmailAddress email : emailList) {
				if(email!=null&&email.getType()!=null&&email.getValue()!=null)
				{
					
					String label = email.getType();
					EmailAddresses = email.getValue();
					if (email.getDisplayName() != null) {
						emailDisplayname = email.getDisplayName();

					}

					MapiContactElectronicAddress emailAddress = new MapiContactElectronicAddress();
					emailAddress.setDisplayName(emailDisplayname);
					emailAddress.setEmailAddress(EmailAddresses);

					if (label.equalsIgnoreCase("home")) {
						ElecAddrPropSet.setEmail1(emailAddress);

					}  if (label.equalsIgnoreCase("work")) {
						ElecAddrPropSet.setEmail2(emailAddress);
					}  if (label.equalsIgnoreCase("other")) {
						ElecAddrPropSet.setEmail3(emailAddress);
					}	

				}
				if(email!=null&&email.getType()==null&&email.getValue()!=null)
				{

					EmailAddresses = email.getValue();
					if (email.getDisplayName() != null) {
						emailDisplayname = email.getDisplayName();

					}

					MapiContactElectronicAddress emailAddress = new MapiContactElectronicAddress();
					emailAddress.setDisplayName(emailDisplayname);
					emailAddress.setEmailAddress(EmailAddresses);
					ElecAddrPropSet.setEmail1(emailAddress);
				}

			}
			}

			contact.setElectronicAddresses(ElecAddrPropSet);
						
				MapiContactPhysicalAddressPropertySet PhysAddrPropSet = new MapiContactPhysicalAddressPropertySet();
				 List<Address> addressList=person.getAddresses();
				 if(addressList!=null)
				 {
				for (Address address : addressList) {

					if (address!= null&&address.getType()!=null) {
						
						String label =address.getType();
						MapiContactPhysicalAddress PhysAddrss = new MapiContactPhysicalAddress();
						
						
						 if (address.getType().equalsIgnoreCase("home")) {
							PhysAddrPropSet.setHomeAddress(PhysAddrss);
						}  if (address.getType().equalsIgnoreCase("work")) {
							PhysAddrPropSet.setWorkAddress(PhysAddrss);
						}  if (address.getType().equalsIgnoreCase("other")) {
							PhysAddrPropSet.setOtherAddress(PhysAddrss);
						}

						if (address.getCity()!=null) {
							PhysAddrss.setCity(address.getCity());
						}
						if (address.getCountry()!=null) {
							PhysAddrss.setCountry(address.getCountry());
						}

						if (address.getPostalCode()!=null) {
							PhysAddrss.setPostalCode(address.getPostalCode());
						}
						if (address.getStreetAddress()!=null) {
							PhysAddrss.setStreet(address.getStreetAddress());
						}
						if (address.getFormattedValue()!=null) {
							PhysAddrss.setAddress(address.getFormattedValue());
						}
						if (address.getPoBox()!=null) {
							PhysAddrss.setPostOfficeBox(address.getPoBox());
						}
						if (address.getRegion()!=null) {
							PhysAddrss.setStateOrProvince(address.getRegion());
						}
						if (address.getCountryCode()!=null) {
							PhysAddrss.setCountryCode(address.getCountryCode());
						}
						

					}
				}
				 }
				 
				contact.setPhysicalAddresses(PhysAddrPropSet);

			
			MapiContactEventPropertySet event = new MapiContactEventPropertySet();					
				List<com.google.api.services.people.v1.model.Event> h = person.getEvents();
				
				if(h!=null)
				{
					for (com.google.api.services.people.v1.model.Event eve : h) {
						String label = eve.getFormattedType();

						if (label.equalsIgnoreCase("Anniversary")) {
							Calendar calendar = Calendar.getInstance();
							com.google.api.services.people.v1.model.Date birthDayDate=eve.getDate();
							calendar.set(birthDayDate.getYear(), birthDayDate.getMonth()-1, birthDayDate.getDay());
							DateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
							String dateEvent = formatter.format(calendar.getTime());

							Date eventDayDate = null;
							try {
								eventDayDate = new SimpleDateFormat("yyyy-MM-dd").parse(dateEvent);
								event.setWeddingAnniversary(eventDayDate);
							} catch (ParseException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
					
						}
					
				}
				}
				

			

				List<Birthday> birthDayList = person.getBirthdays();	
				if(birthDayList!=null)
				{
					
				for (Birthday birthday : birthDayList) {
					
					if(birthday!=null)
					{
						Calendar calendar = Calendar.getInstance();
						com.google.api.services.people.v1.model.Date birthDayDate=birthday.getDate();
						calendar.set(birthDayDate.getYear(), birthDayDate.getMonth()-1, birthDayDate.getDay());
						DateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
						String dateEvent = formatter.format(calendar.getTime());

						Date eventDayDate = null;
						try {
							eventDayDate = new SimpleDateFormat("yyyy-MM-dd").parse(dateEvent);
							event.setBirthday(eventDayDate);
						} catch (ParseException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
				
			      		

				}
					
				}
		     }


			if (event != null) {
				contact.setEvents(event);
			}

			MapiContactProfessionalPropertySet ProfPropSet = new MapiContactProfessionalPropertySet();
		
				List<Organization> org = person.getOrganizations();
				if(org!=null)
				{
				for (Organization organization : org) {
				
						if (organization.getDepartment()!=null) {
							ProfPropSet.setDepartmentName(organization.getDepartment());
					      }
						if (organization.getName()!=null) {
							ProfPropSet.setCompanyName(organization.getName());
					      }
						if (organization.getTitle()!=null) {
							ProfPropSet.setTitle(organization.getTitle());
					      }
						if (organization.getJobDescription()!=null) {
							ProfPropSet.setProfession(organization.getJobDescription());
					      }
						
					}

				}

				List<Occupation> OccupationList=person.getOccupations();
				if(OccupationList!=null)
				{
				for (Occupation Occupation : OccupationList) {
					
					
						if(Occupation.getValue()!=null)
						{
							ProfPropSet.setProfession(Occupation.getValue());

						}
					
					}
					
				}

					contact.setProfessionalInfo(ProfPropSet);

					MapiContactPersonalInfoPropertySet info = new MapiContactPersonalInfoPropertySet();
					List<Url> urleList=person.getUrls();
					if(urleList!=null)
					{
						for (Url url : urleList) {
							if(url!=null)
							{
								info.setPersonalHomePage(url.getValue());
							}
						}
					}
					
					
					List<Gender> genderList=person.getGenders();
					if(genderList!=null)
					{
						for (Gender Gender : genderList) {
							// info.setGender(Gender.getFormattedValue();
						}
					}
					

					if (person.getSkills() != null) {
						List<Skill> hobby = person.getSkills();
						for (Skill skill : hobby) {
							
							info.setHobbies(skill.getValue());
							
						}

					}
					
			

//					try {
//						if (person.get != null) {
//
//							info.setNotes(entry.getPlainTextContent());
//						//	System.out.println(entry.getPlainTextContent());
//
//						}
//
//					} catch (Exception e) {
//						// TODO: handle exception
//						e.printStackTrace();
//					}

					if (info != null) {
						contact.setPersonalInfo(info);
					}

					if (fullNameToDisplay != null) {
						contact.save(EmailWizardApplication.detinationPath + File.separator + count + "-" + fullNameToDisplay + ".vcf",
								ContactSaveFormat.VCard);
//						
//						FileStream fs = new FileStream(GoogleMaineFrame.detinationPath + File.separator + i + "-" + fullNameToDisplay + ".vcf", FileMode.Open);
//
//						contact= MapiContact.fromVCard(fs, StandardCharsets.UTF_8);
//						contact.save(GoogleMaineFrame.detinationPath + File.separator + i + "-" + fullNameToDisplay + ".vcf",
//								ContactSaveFormat.VCard);
					} else {
						contact.save(EmailWizardApplication.detinationPath+ File.separator + count + "-" + EmailAddresses + ".vcf",
								ContactSaveFormat.VCard);
						
						
					}

					EmailWizardApplication.downloadingFileName.setText(fullNameToDisplay+"_"+count);
		            EmailWizardApplication.modelDownloading.setValueAt(count, EmailWizardApplication.rownCount, 2);
		            
		            int prog = (count * 100) /  connections.size();
					EmailWizardApplication.progressBar_Downloading.setValue(prog);
			     	count++;

				} catch (Exception e) {
					// TODO: handle exception
					e.printStackTrace();

					StringWriter errors = new StringWriter();
					e.printStackTrace(new PrintWriter(errors));

					
				}
		            }
		        } else {
		            System.out.println("No connections found.");
		        } 
				



		} catch (Exception ep) {

			StringWriter errors = new StringWriter();
			ep.printStackTrace(new PrintWriter(errors));
			EmailWizardApplication.logger.warn(
					errors + System.lineSeparator());

            ep.printStackTrace();
			if (ep.getMessage().contains("www.google.com") 
					|| ep.getMessage().contains("oauth2.googleapis.com")
					|| ep.getMessage().contains("No route to host: connect")
					||ep.getMessage().contains("Failed to refresh access token: Connection reset")
					||ep.getMessage().contains("Connection reset")
					||ep.getMessage().contains("Software caused connection abort: connect")) {

			ep.printStackTrace();
			
			EmailWizardApplication.lblNoInternetConnection.setVisible(true);
			System.out.println("No Internet Connection");
				while (!checkInternet()) { 						
				}
			EmailWizardApplication.lblNoInternetConnection.setVisible(false);

			downoadContacts();


			}

		} 

	}
	private static HttpRequestInitializer setHttpTimeout(final HttpRequestInitializer requestInitializer) {
        return new HttpRequestInitializer() {
            @Override
            public void initialize(HttpRequest httpRequest) throws IOException {
                requestInitializer.initialize(httpRequest);
                httpRequest.setConnectTimeout(3 * 60000);  // 3 minutes connect timeout
                httpRequest.setReadTimeout(3 * 60000);  // 3 minutes read timeout
            }

	
        };
    }
	
	public static boolean checkInternet() {
		try {
			URL url = new URL("http://www.google.com");
			URLConnection connection = url.openConnection();
			connection.connect();
			System.out.println("Internet is connected");
			return true;
		} catch (MalformedURLException e) {
			System.out.println("Internet is not connected");
		} catch (IOException e) {
			System.out.println("Internet is not connected");
		}
		return false;
	}
	

}
