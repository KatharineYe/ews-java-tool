package com.ewslib.tools;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URI;
import java.util.Calendar;
import java.util.Date;
import microsoft.exchange.webservices.data.ConflictResolutionMode;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.ExtendedPropertyDefinition;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderView;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.ItemView;
import microsoft.exchange.webservices.data.MapiPropertyType;
import microsoft.exchange.webservices.data.MessageBody;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.WellKnownFolderName;


class Email implements Runnable {
	private String userName, password, domain, recipient;
	private int emailCount;
	
	private void setUserName(String userName) {
		this.userName = userName;
	}
	
	private void setPassword(String password) {
		this.password = password;
	}
	
	private void setDomain(String domain) {
		this.domain = domain;
	}
	
	private void setRecipient(String recipient) {
		this.recipient = recipient;
	}
	
	private void setEmailCount(int emailCount) {
		this.emailCount = emailCount;
	}
	
	private void splitEmailAddress(String emailAddress) {
		String[] email = emailAddress.split("@");
		this.setUserName(email[0]);
		this.setDomain(email[1]);
	}
	
	public String getUserName() {
		return userName;
	}

	public String getPassword() {
		return password;
	}

	public String getDomain() {
		return domain;
	}

	public String getRecipient() {
		return recipient;
	}

	public int getEmailCount() {
		return emailCount;
	}

	
	public void operation(BufferedReader bfReader) {
		String input = null;
		int selectAction = 0;
		try {
			selectAction = Integer.parseInt(bfReader.readLine());
		} catch (NumberFormatException e) {
			selectAction = -1;
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		switch(selectAction) {
		case 1: {
			System.out.println("已选择邮件操作：[1-发送单封邮件]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"收件人邮件地址\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd user02@mail1.com");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 3) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String recipient = parameters[2].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				boolean isRecipientFormatCorrect = Format.emailFormatCheck(parameters[2]);
				if (isEmailAddressFormatCorrect && isRecipientFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.sendEmail(userName, password, domain, recipient);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 2: {
			System.out.println("已选择邮件操作：[2-发送多封邮件]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"收件人邮件地址\"，\"发送邮件数量\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd user02@mail1.com 10");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 4) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String recipient = parameters[2].trim();
				String emailCount = parameters[3].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				boolean isRecipientFormatCorrect = Format.emailFormatCheck(recipient);
				boolean isInt = Format.numberCheck(emailCount);
				if (isEmailAddressFormatCorrect && isRecipientFormatCorrect && isInt) {
					splitEmailAddress(emailAddress);
					int eCount = Integer.parseInt(emailCount);
					this.sendEmails(userName, password, domain, recipient, eCount);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 3: {
			System.out.println("已选择邮件操作：[3-选择从收件箱、草稿箱、已发送文件夹中删除邮件]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"选择删除邮件的文件夹\"，以空格符分隔并以换行符结束！");
			System.out.println("可选文件夹：收件箱（inbox），草稿箱（drafts），已发送（sentitems）");
			System.out.println("例如：user01@mail1.com pwd inbox");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 3) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String eFolder = parameters[2].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.deleteAllEmailInFolder(userName, password, domain, eFolder);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 4: {
			System.out.println("已选择邮件操作：[4-清空已删除文件夹中的邮件]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 2) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.permanentDeleteAllEmailInTrash(userName, password, domain);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		default: {
			System.out.println("错误：输入无效！");
			break;}
		}
	}
	
	
/**
 * Call this method to create an email by using account.
 * @param userName User name of specified account.
 * @param password Password of specified account.
 * @param domain Domain of specified account.
 * @param date A Date for specified date.
 */
	private void sendEmail(String userName, String password, String domain, String recipient) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			EmailMessage msg = new EmailMessage(service);
			msg.setSubject("Auto email.");
			msg.setBody(MessageBody.getMessageBodyFromText("Sent by EWS."));
			msg.getToRecipients().add(recipient);
			msg.sendAndSaveCopy();
			System.out.println("Sent email successfully.");
			Thread.sleep(100);
		} catch (Exception e) {
			System.out.println("Error: Fail to send email!");
			e.printStackTrace();
		}
	}
	
	
/**
 * Call this method to create email by using account.
 * @param userName User name of specified account.
 * @param password Password of specified account.
 * @param domain Domain of specified account.
 * @param date A Date for specified date.
 */
	private void createEmails(String userName, String password, String domain, String recipient) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			int i = 0;
			int count = this.getEmailCount() + 1;
			String eSubject = null;
			while (emailCount > 1) {
				EmailMessage msg = new EmailMessage(service);
				synchronized (this) {
					msg.getToRecipients().add(recipient);
					i = count - emailCount;
					eSubject = "Auto email " + i + ".";
					msg.setSubject(eSubject);
					msg.setBody(MessageBody.getMessageBodyFromText("Sent by EWS."));
					msg.sendAndSaveCopy();
//					System.out.println(Thread.currentThread().getName() + "->" + i + ": Successfully sent email " + "[" + eSubject +"].");
					System.out.println(i + ": Successfully sent email " + "[" + eSubject +"].");
					Thread.sleep(200);
					emailCount--;
				}
			}
		} catch (Exception e) {
			System.out.println("Error: Fail to send email!");
			e.printStackTrace();
		}
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		createEmails(userName, password, domain, recipient);
	}
	
	private void sendEmails(String userName, String password, String domain, String recipient, int count) {
		this.setUserName(userName);
		this.setPassword(password);
		this.setDomain(domain);
		this.setRecipient(recipient);
		this.setEmailCount(count);
		Thread thread1 = new Thread(this, "Thread 1");
		Thread thread2 =new Thread(this, "Thread 2");
		thread1.start();
		thread2.start();
		while(thread1.isAlive() || thread2.isAlive()) {
			try {
				Thread.sleep(1000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}

	
/**
 * Call this method to delete all email in specified email root folder.
 * @param userName User name of specified account.
 * @param password Password of specified account.
 * @param domain Domain of specified account.
 */
	private void deleteAllEmailInFolder(String userName, String password, String domain, String folderName) {
		WellKnownFolderName fName = null;
		if (folderName.equals("inbox")) {
			fName = WellKnownFolderName.Inbox;
		} else if (folderName.equals("sentitems")) {
			fName = WellKnownFolderName.SentItems;
		} else if (folderName.equals("drafts")) {
			fName = WellKnownFolderName.Drafts;
		}
		try {
			EWS ews = new EWS();
			ExchangeService service = ews.connectEWS(userName, password, domain);
			Folder folder = Folder.bind(service, fName); 
			int eCount = folder.getTotalCount();
			if (eCount > 0) {
				System.out.println("Total " + eCount +" email in " + folderName + "!");
				folder.empty(DeleteMode.MoveToDeletedItems, true);
				System.out.println("All email in " + folderName + " have been deleted!");
			} else if (eCount == 0) {
				System.out.println("No email in " + folderName + "!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
		
/**
 * Call this method to delete email in Trash.
 * @param userName User name of specified account.
 * @param password Password of specified account.
 * @param domain Domain of specified account.
 */
	private void permanentDeleteAllEmailInTrash(String userName, String password, String domain) {
			try {
				EWS ews = new EWS();
				ExchangeService service = ews.connectEWS(userName, password, domain);
				Folder delete = Folder.bind(service, WellKnownFolderName.DeletedItems); 
				int eCount = delete.getTotalCount();
				if (eCount > 0) {
					System.out.println("Total " + eCount +" email in Trash!");
					delete.empty(DeleteMode.HardDelete, true);
					System.out.println("All email in Trash have been deleted!");
				} else if (eCount == 0) {
					System.out.println("No email in Trash folder!");
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
	}
	
	
	public void setEmailFlag(String userName, String password, String domain, String emailsubject) {
		try {
			ExtendedPropertyDefinition PidTagFlagStatus_e = new ExtendedPropertyDefinition(0x1090, MapiPropertyType.Integer);
			ExtendedPropertyDefinition PidTagFlagStatus_t = new ExtendedPropertyDefinition(0x1095, MapiPropertyType.Integer);
			EWS ews = new EWS();
			ExchangeService service = ews.connectEWS(userName, password, domain);
			Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox); 
			ItemView view = new ItemView(300);
			FindItemsResults<Item> findResults = service.findItems(inbox.getId(), view);
			int i = 0;
			for (Item item : findResults.getItems()) {
				i++;
				EmailMessage message = EmailMessage.bind(service, item.getId());
				String subject = message.getSubject();
				if (subject.equals(emailsubject)) {
System.out.println(i + " " + subject);
					message.setExtendedProperty(PidTagFlagStatus_e, 2); //2 - Flag; 1 - Completed; 0 - Clear;
					message.setExtendedProperty(PidTagFlagStatus_t, 6); 
					message.update(ConflictResolutionMode.AutoResolve);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	public boolean getEmailFlag(String userName, String password, String domain, String emailsubject) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		boolean flag = false;
		try {

			Folder rootfolder = Folder.bind(service, WellKnownFolderName.Root);
			for (Folder folder : rootfolder.findFolders(new FolderView(100))) {
				if(folder.getDisplayName().equals("To-Do Search")) {
					ItemView view = new ItemView(300);
					FindItemsResults<Item> findResults = service.findItems(folder.getId(), view);
					int i = 0;
					for (Item item : findResults.getItems()) {
						String subject = item.getSubject();
						if (subject.equals(emailsubject)) {
							System.out.println(item.getSubject());
							flag = true;
						} else {
							flag = false;
						}
					}
				}
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return flag;
	}
}
