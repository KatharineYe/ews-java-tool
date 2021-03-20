package com.ewslib.tools;

import java.net.URI;
import java.util.Scanner;
import java.util.TimeZone;
import java.io.IOException;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.ExchangeCredentials;


public class EWS {
	public String uri_mail1 = "https://10.1.15.xxx/EWS/Exchange.asmx";
	public String uri_mail2 = "https://10.1.15.xxx/EWS/Exchange.asmx";
	public String uri_mail3 = "https://10.1.15.xxx/EWS/Exchange.asmx";
	
	
	ExchangeService connectEWS(String userName, String password, String domain) {
		ExchangeService service = null;
		try {
			service = new ExchangeService(ExchangeVersion.Exchange2010_SP2, TimeZone.getTimeZone("GMT+8"));
			ExchangeCredentials credentials = new WebCredentials(userName, password, domain);
			service.setCredentials(credentials);
			String uri = null;
			if (domain.equals("mail1.com")) {
				uri = uri_mail1;
			} else if (domain.equals("mail2.com")) {
				uri = uri_mail2;
			} else if (domain.equals("mail3.com")) {
				uri = uri_mail3;
			} 
			service.setUrl(new URI(uri));
		} catch (Exception e) {
			e.printStackTrace();
		}	
		return service;
	}
	
	
	public static void main(String[] args) {
		System.out.println("******************************************");
		System.out.println("*** Easy Use EWS Tool  ***");
		System.out.println("*** Version: 2.0 ***");
		System.out.println("*** Author: Capital_K ***");
		System.out.println("******************************************");
		BufferedReader bfReader = new BufferedReader(new InputStreamReader(System.in));
		while (true) {
			System.out.println();
			System.out.println("请选择要操作的对象：1-邮件 2-邮件文件夹 3-日历 0-退出");
			int targetNum = -1;
			try {
				targetNum = Integer.parseInt(bfReader.readLine());
			} catch (NumberFormatException e) {
				targetNum = -1;
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			TimeZone tz = TimeZone.getDefault();
			switch(targetNum) {
			case 0:
				try {
					bfReader.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
				System.exit(0);
			case 1:
				Email email = new Email();
				System.out.println("已选择操作对象 [1-邮件]");
				System.out.println("邮件可选操作：1-发送单封邮件 2-发送多封邮件 3-选择从收件箱、草稿箱、已发送文件夹中删除邮件 4-清空已删除文件夹中的邮件");
				email.operation(bfReader);
				break;
			case 2:
				EmailFolder emailFolder = new EmailFolder();
				System.out.println("已选择操作对象 [2-邮件文件夹]");
				System.out.println("请选择可选操作：1-创建邮件文件夹 2-删除单个邮件文件夹 3-删除多个邮件文件夹 4-重命名邮件文件夹 5-移动邮件文件夹");
				emailFolder.operation(bfReader);
				break;
			case 3:
				Calendar calendar = new Calendar();
				System.out.println("已选择操作对象 [3-日历]");
				System.out.println("请选择可选操作：0-删除全部约会和会议 1-创建单一约会 2-创建重复约会 3-创建单一会议 4-创建重复会议");
//				calendar.operation(bfReader);
//				calendar.createSingleAllDayAppointment("username", "password", "mail3.com", "all-day test", "tj", "", "", "");
				TimeZone.setDefault(tz);
				break;
			default:
				System.out.println("错误：输入无效！！！");
				break;
			} 
		}
		
	}
}
