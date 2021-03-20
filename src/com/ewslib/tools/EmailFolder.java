package com.ewslib.tools;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URI;
import java.util.Calendar;
import java.util.Date;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.FolderPermissionCollection;
import microsoft.exchange.webservices.data.FolderView;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.ItemView;
import microsoft.exchange.webservices.data.Mailbox;
import microsoft.exchange.webservices.data.ManagedFolderInformation;
import microsoft.exchange.webservices.data.MessageBody;
import microsoft.exchange.webservices.data.SearchFilter;
import microsoft.exchange.webservices.data.ServiceResponseException;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.WellKnownFolderName;

public class EmailFolder {
	public String FOLDER_DELETE = "Deleted Items";
    public String FOLDER_SENT = "Sent Items";
    public String FOLDER_OUTBOX = "Outbox";
    public String FOLDER_JUNK = "Junk E-Mail";
    public String FOLDER_DRAFT = "Drafts";
    public String FOLDER_INBOX = "Inbox";
    private String userName, password, domain;
    
    private void setUserName(String userName) {
		this.userName = userName;
	}
	
	private void setPassword(String password) {
		this.password = password;
	}
	
	private void setDomain(String domain) {
		this.domain = domain;
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
			System.out.println("已选择邮件文件夹操作：[1-创建邮件文件夹]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"新建文件夹\"，\"文件夹级别\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd NewFolder01 1");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 4) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String folderName = parameters[2].trim();
				String folderLevel = parameters[3].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				boolean isInt = Format.numberCheck(folderLevel);
				if (isEmailAddressFormatCorrect && isInt) {
					splitEmailAddress(emailAddress);
					int level = Integer.parseInt(folderLevel);
					this.createFolder(userName, password, domain, folderName, level);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 2: {
			System.out.println("已选择邮件文件夹操作：[2-删除单个邮件文件夹]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"删除文件夹\"，\"文件夹级别\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd NewFolder01 1");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 4) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String folderName = parameters[2].trim();
				String folderLevel = parameters[3].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				boolean isInt = Format.numberCheck(folderLevel);
				if (isEmailAddressFormatCorrect && isInt) {
					splitEmailAddress(emailAddress);
					int level = Integer.parseInt(folderLevel);
					this.deleteFolder(userName, password, domain, folderName, level);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 3: {
			System.out.println("已选择邮件文件夹操作：[3-删除多个邮件文件夹]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"文件夹级别\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd 1");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 3) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String folderLevel = parameters[2].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				boolean isInt = Format.numberCheck(folderLevel);
				if (isEmailAddressFormatCorrect && isInt) {
					splitEmailAddress(emailAddress);
					int level = Integer.parseInt(folderLevel);
					this.deleteFolders(userName, password, domain, level);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 4: {
			System.out.println("已选择邮件文件夹操作：[4-重命名邮件文件夹]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"原文件夹名称\"，\"新文件夹名称\"，\"文件夹级别\"，以空格符分隔并以换行符结束！");
			System.out.println("例如：user01@mail1.com pwd feedback Customer-Feedback 2");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 5) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String orgFolder = parameters[2].trim();
				String updateFolder = parameters[3].trim();
				String folderLevel = parameters[4].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				boolean isInt = Format.numberCheck(folderLevel);
				if (isEmailAddressFormatCorrect && isInt) {
					splitEmailAddress(emailAddress);
					int level = Integer.parseInt(folderLevel);
					this.renameFolder(userName, password, domain, orgFolder, updateFolder, level);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 5: {
			System.out.println("已选择邮件文件夹操作：[5-移动邮件文件夹]");
			System.out.println("请输入\"发件箱邮件地址\"，\"发件箱密码\"，\"文件夹名称\"，\"原位置\"，\"目标位置\"，以空格符分隔并以换行符结束！");
			System.out.println("可选位置：根目录（root），收件箱（inbox）");
			System.out.println("例如：user01@mail1.com pwd A1 root inbox");
			try {
				input = bfReader.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(" ");
			if (parameters.length == 5) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String folderName = parameters[2].trim();
				String orgLoc = parameters[3].trim();
				String destLoc = parameters[4].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.moveFolder(userName, password, domain, folderName, orgLoc, destLoc);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		default:
			System.out.println("错误：输入无效！");
			break;
		}
    }
    
    
	public void createFolder(String userName, String password, String domain, String folderName, int level) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			WellKnownFolderName fn = null;
			if (level == 1) {
				fn = WellKnownFolderName.MsgFolderRoot;
			} else if (level == 2) {
				fn = WellKnownFolderName.Inbox;
			}
			Folder rootfolder = Folder.bind(service, fn);
			int fCount = rootfolder.getChildFolderCount();
			for (Folder folder : rootfolder.findFolders(new FolderView(fCount))) {
				String name = folder.getDisplayName();
				if (name.equals(folderName)) {
					System.out.println("Folder \"" + folderName + "\" is exist!");
					return;
				} 
			}
			Folder nFolder = new Folder(service);
			nFolder.setDisplayName(folderName);
			if (level == 1) {
				nFolder.save(WellKnownFolderName.MsgFolderRoot);
				System.out.println("Create email root folder - \"" + folderName + "\"");
			} else if (level == 2) {
				nFolder.save(WellKnownFolderName.Inbox);
				System.out.println("Create email subfolder - \"" + folderName + "\"");
			} 
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	public void deleteFolder(String userName, String password, String domain, String folderName, int level) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			WellKnownFolderName fn = null;
			if (level == 1) {
				fn = WellKnownFolderName.MsgFolderRoot;
			} else if (level == 2) {
				fn = WellKnownFolderName.Inbox;
			}
			Folder rootfolder = Folder.bind(service, fn);
			int fCount = rootfolder.getChildFolderCount();
			boolean isExist = false;
			for (Folder folder : rootfolder.findFolders(new FolderView(fCount))) {
				String name = folder.getDisplayName();
				if (name.equals(folderName)) {
					isExist = true;
					folder.delete(DeleteMode.HardDelete);
					System.out.println("Folder \"" + name + "\" is deleted!");
					break;
				} 
			}
			if (isExist == false) {
				System.out.println("Error: Folder \"" + folderName + "\" is not exist, please recheck!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	public void deleteFolders(String userName, String password, String domain, int level) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
        String name = null;
        try {
            WellKnownFolderName fn = null;
            if (level == 1) {
                fn = WellKnownFolderName.MsgFolderRoot;
            } else if (level == 2) {
                fn = WellKnownFolderName.Inbox;
            }
            Folder rootfolder = Folder.bind(service, fn);
            int fCount = rootfolder.getChildFolderCount();
            for (Folder folder : rootfolder.findFolders(new FolderView(fCount))) {
                if (folder.getFolderClass() == null) {
                    name = folder.getDisplayName();
                    if (name.equals(this.FOLDER_INBOX) || name.equals(this.FOLDER_DRAFT) || name.equals(this.FOLDER_OUTBOX) || name.equals(this.FOLDER_SENT) || name.equals(this.FOLDER_DELETE) || name.equals(this.FOLDER_JUNK)) {
                        break;
                    } else {
                        folder.delete(DeleteMode.HardDelete);
                        System.out.println("Folder \"" + name + "\" is deleted!");
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Error: Failed to delete email folder \"" + name + "\"!");
            e.printStackTrace();
        }
    }
	
	
	public void renameFolder(String userName, String password, String domain, String orgFolderName, String updateFolderName, int level) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			WellKnownFolderName fn = null;
			if (level == 1) {
				fn = WellKnownFolderName.MsgFolderRoot;
			} else if (level == 2) {
				fn = WellKnownFolderName.Inbox;
			}
			Folder rootfolder = Folder.bind(service, fn);
			int fCount = rootfolder.getChildFolderCount();
			boolean isExist = false;
			for (Folder folder : rootfolder.findFolders(new FolderView(fCount))) {
				String name = folder.getDisplayName();
				if (name.equals(orgFolderName)) {
					isExist = true;
					folder.setDisplayName(updateFolderName);
					folder.update();
					System.out.println("Update folder from \"" + orgFolderName + "\" to \"" + updateFolderName + "\"!");
					break;
				}
			}
			if (!isExist) {
				System.out.println("Error: Folder \"" + orgFolderName + "\" is not exist, please recheck!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	
	public void moveFolder(String userName, String password, String domain, String folderName, String orig, String dest) {
		WellKnownFolderName orf = null;
		WellKnownFolderName drf = null;
		if (orig.toLowerCase().equals("root")) {
			orf = WellKnownFolderName.MsgFolderRoot;
		} else if (orig.toLowerCase().equals("inbox")) {
			orf = WellKnownFolderName.Inbox;
		}
		if (dest.toLowerCase().equals("root")) {
			drf = WellKnownFolderName.MsgFolderRoot;
		} else if (dest.toLowerCase().equals("inbox")) {
			drf = WellKnownFolderName.Inbox;
		}
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			Folder rootfolder = Folder.bind(service, orf);
			int fCount = rootfolder.getChildFolderCount();
			boolean isExist = false;
			for (Folder folder : rootfolder.findFolders(new FolderView(fCount))) {
				String name = folder.getDisplayName();
				if (name.equals(folderName)) {
					isExist = true;
					folder.move(drf);
					System.out.println("Folder \"" + folderName + "\" is moved");
					break;
				}
			}
			if (!isExist) {
				System.out.println("Error: Folder \"" + folderName + "\" is not exist, please recheck!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
