package com.ewslib.tools;

import java.io.File;
import java.io.IOException;
import java.io.BufferedReader;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.Locale;
import java.util.HashMap;
import java.util.List;
import java.util.TimeZone;
import java.util.Collection;
import java.util.Locale.Category;
import microsoft.exchange.webservices.data.*;


public class Calendar {
	private String userName, password, domain, recipient;
	
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
	
	/**
	 * Appointment:创建单一约会
	 */
	public void createSingleAppointment(String userName, String password, String domain, String subject, String location, String startsTime, String endsTime, String body) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			Appointment appointment = new Appointment(service);
			appointment.setSubject(subject); // 约会标题
			appointment.setLocation(location); // 约会地点
			appointment.setBody(MessageBody.getMessageBodyFromText(body)); // 约会正文
			HashMap times = getStartsAndEndsTime(startsTime, endsTime);
			appointment.setStart((Date) times.get("startsTime"));
			appointment.setEnd((Date) times.get("endsTime")); 
			appointment.save(WellKnownFolderName.Calendar, SendInvitationsMode.SendToNone);
		} catch (Exception e) {
			System.out.println("Error: Fail to create appointment!");
			e.printStackTrace();
		}
	}
	
	/**
	 * Appointment:创建单一会议
	 */
	public void createSingleMeeting(String userName, String password, String domain, String required, String optional, String subject, String location, String startsTime, String endsTime, String body) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		String[] requiredList, optionalList;
		try {
			Appointment appointment = new Appointment(service);
			if (!required.equals("") && !required.equals(null)) {
				requiredList = required.split(";");
				for (String r : requiredList) {
					appointment.getRequiredAttendees().add(r);
				}
			} 
			if (!optional.equals("") && !optional.equals(null)) {
				optionalList = optional.split(";");
				for (String o : optionalList) {
					appointment.getOptionalAttendees().add(o);
				}
			}
			if ((required.equals("") || required.equals(null)) && (optional.equals("") || optional.equals(null))) {
				System.out.println("输入无效！");
				return;
			}
			appointment.setSubject(subject); // 约会标题
			appointment.setLocation(location); // 约会地点
			appointment.setBody(MessageBody.getMessageBodyFromText(body)); // 约会正文
			HashMap times = getStartsAndEndsTime(startsTime, endsTime);
			appointment.setStart((Date) times.get("startsTime"));
			appointment.setEnd((Date) times.get("endsTime")); 
			appointment.save(WellKnownFolderName.Calendar);
		} catch (Exception e) {
			System.out.println("Error: Fail to create appointment!");
			e.printStackTrace();
		}
	}

	/**
	 * Meeting:创建重复约会
	 */
	public void createRecurringAppointment(String userName, String password, String domain, String subject, String location, String startsTime, String endsTime, String recurrenceType, String recurrenceEndsDate, String body) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		try {
			Appointment appointment = new Appointment(service);
			appointment.setSubject(subject); // 约会标题
			appointment.setLocation(location); // 约会地点
			appointment.setBody(MessageBody.getMessageBodyFromText(body)); // 约会正文
			// 约会开始时间和结束时间
			HashMap times = getStartsAndEndsTime(startsTime, endsTime);
			appointment.setStart((Date) times.get("startsTime"));
			appointment.setEnd((Date) times.get("endsTime")); 
			// 约会循环类型
			Date dateRecurrenceStarts = appointment.getStart();
			Date dateRecurrenceEnds = getRecurrenceEndsDate(recurrenceEndsDate);
			Recurrence recurrences = null;
			switch(recurrenceType.toLowerCase()) {
			case "daily":
				recurrences = new Recurrence.DailyPattern(dateRecurrenceStarts, 1); // 默认interval为1天
				break;
			case "weekly":
				/**
				 * @param 重复会议开始日期
				 * @param 重复间隔为一周一次
				 * @param 重复会议发生在每周星期一
				 */
				recurrences = new Recurrence.WeeklyPattern(dateRecurrenceStarts, 1, DayOfTheWeek.Monday); // 默认每周一
				break;
			case "monthly":
				/**
				 * @param 重复会议开始日期
				 * @param 重复间隔为一月一次
				 * @param 重复会议发生在每月一号
				 */
				recurrences = new Recurrence.MonthlyPattern(dateRecurrenceStarts, 1, 1); // 默认每月一号
				break;
			case "yearly":
				/**
				 * @param 重复会议开始日期
				 * @param 重复间隔为一月一次
				 * @param 重复会议发生在每月一号
				 */
				recurrences = new Recurrence.YearlyPattern(dateRecurrenceStarts, Month.June, 1); // 默认每月一号
				break;
			default:
				System.out.println("输入无效！");
				break;
			}
			appointment.setRecurrence(recurrences);
			appointment.getRecurrence().setStartDate(dateRecurrenceStarts);
			appointment.getRecurrence().setEndDate(dateRecurrenceEnds);
			appointment.save(WellKnownFolderName.Calendar, SendInvitationsMode.SendToNone);
		} catch (Exception e) {
			System.out.println("Error: Fail to create appointment!");
			e.printStackTrace();
		}
	}
	
	/**
	 * Meeting:创建重复会议
	 */
	public void createRecurringMeeting(String userName, String password, String domain, String required, String optional, String subject, String location, String startsTime, String endsTime, String recurrenceType, String recurrenceEndsDate, String body) {
		EWS ews = new EWS();
		ExchangeService service = ews.connectEWS(userName, password, domain);
		String[] requiredList, optionalList;
		try {
			Appointment appointment = new Appointment(service);
			if (!required.equals("") && !required.equals(null)) {
				requiredList = required.split(";");
				for (String r : requiredList) {
					appointment.getRequiredAttendees().add(r);
				}
			} 
			if (!optional.equals("") && !optional.equals(null)) {
				optionalList = optional.split(";");
				for (String o : optionalList) {
					appointment.getOptionalAttendees().add(o);
				}
			}
			if ((required.equals("") || required.equals(null)) && (optional.equals("") || optional.equals(null))) {
				System.out.println("输入无效！");
				return;
			}
			appointment.setSubject(subject); // 约会标题
			appointment.setLocation(location); // 约会地点
			appointment.setBody(MessageBody.getMessageBodyFromText(body)); // 约会正文
			// 约会开始时间和结束时间
			HashMap times = getStartsAndEndsTime(startsTime, endsTime);
			appointment.setStart((Date) times.get("startsTime"));
			appointment.setEnd((Date) times.get("endsTime")); 
			// 约会循环类型
			Date dateRecurrenceStarts = appointment.getStart();
			Date dateRecurrenceEnds = getRecurrenceEndsDate(recurrenceEndsDate);
			Recurrence recurrences = null;
			switch(recurrenceType.toLowerCase()) {
			case "daily":
				recurrences = new Recurrence.DailyPattern(dateRecurrenceStarts, 1); // 默认interval为1天
				break;
			case "weekly":
				/**
				 * @param 重复会议开始日期
				 * @param 重复间隔为一周一次
				 * @param 重复会议发生在每周星期一
				 */
				recurrences = new Recurrence.WeeklyPattern(dateRecurrenceStarts, 1, DayOfTheWeek.Monday); // 默认每周一
				break;
			case "monthly":
				/**
				 * @param 重复会议开始日期
				 * @param 重复间隔为一月一次
				 * @param 重复会议发生在每月一号
				 */
				recurrences = new Recurrence.MonthlyPattern(dateRecurrenceStarts, 1, 1); // 默认每月一号
				break;
			case "yearly":
				/**
				 * @param 重复会议开始日期
				 * @param 重复间隔为一月一次
				 * @param 重复会议发生在每月一号
				 */
				recurrences = new Recurrence.YearlyPattern(dateRecurrenceStarts, Month.June, 1); // 默认每月一号
				break;
			default:
				System.out.println("输入无效！");
				break;
			}
			appointment.setRecurrence(recurrences);
			appointment.getRecurrence().setStartDate(dateRecurrenceStarts);
			appointment.getRecurrence().setEndDate(dateRecurrenceEnds);
			appointment.save(WellKnownFolderName.Calendar);
		} catch (Exception e) {
			System.out.println("Error: Fail to create appointment!");
			e.printStackTrace();
		}
	}
	
	/**
	 * 删除全部日程
	 */
	public void deleteAllEvents(String userName, String password, String domain) {
		WellKnownFolderName folderName = WellKnownFolderName.Calendar;
		try {
			ExchangeService service = new EWS().connectEWS(userName, password, domain);
			int curYear = Util.getCurrYear();
			// 检索
			for (int i=0; i<=30; i++) {
				System.out.println("检索时间: " + curYear + " 年...");
				CalendarView view = new CalendarView(Util.getYearFirst(curYear), Util.getYearLast(curYear));
				FindItemsResults<Appointment> findResults = service.findAppointments(folderName, view);
				for (Appointment appointment : findResults.getItems()) {
					System.out.println("Get " + (appointment.getIsRecurring() ? "Recurring" : "Single") + " event: [" + appointment.getSubject() + "]" + appointment.getIsAllDayEvent());
					if (appointment.getIsRecurring()) {
						PropertySet propSet = new PropertySet(
								AppointmentSchema.Subject,
	                            AppointmentSchema.Location,
	                            AppointmentSchema.Start, 
	                            AppointmentSchema.End,
	                            AppointmentSchema.AppointmentType,
	                            AppointmentSchema.Id,
	                            AppointmentSchema.DeletedOccurrences
								);
						Appointment.bindToRecurringMaster(service, appointment.getId(), propSet);
					} 
					if (!appointment.equals(null)) {
						appointment.delete(DeleteMode.HardDelete);
					}
				}
				curYear++;
			}			
		} catch (Exception e) {
			e.printStackTrace();
		}
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
		case 0: {
			System.out.println("已选择日历操作：[0-删除全部约会和会议]");
			System.out.println("请输入\"邮件地址\"，\"邮箱密码\"，以逗号分隔并以换行符结束！");
			System.out.println("例如：usr01@mail1.com, pwd");
			try {
				input = bfReader.readLine();
				if (input.equals("") || input.equals(null)) {
					throw new IOException("输入无效！");
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(",");
			if (parameters.length == 2) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.deleteAllEvents(userName, password, domain);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 1: {
			System.out.println("已选择日历操作：[1-创建单一约会]");
			System.out.println("请输入\"邮件地址\"，\"邮箱密码\"，\"约会主题\"，\"约会地点\"，\"约会开始时间(可为空)\"，\"约会结束时间(可为空)\"，\"约会正文\"，以逗号分隔并以换行符结束！");
			System.out.println("例如：usr01@mail1.com,pwd,Single-Appointment,Tianjin,2018-04-10 10:00:00,2018-04-10 11:00:00,Appointment-Body.");
			try {
				input = bfReader.readLine();
				if (input.equals("") || input.equals(null)) {
					throw new IOException("输入无效！");
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(",");
			if (parameters.length == 7) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String subject = parameters[2].trim();
				String location = parameters[3].trim();
				String startsTime = parameters[4].trim();
				String endsTime = parameters[5].trim();
				String body = parameters[6].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.createSingleAppointment(userName, password, domain, subject, location, startsTime, endsTime, body);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 2: {
			System.out.println("已选择日历操作：[2-创建重复约会]");
			System.out.println("请输入\"邮件地址\"，\"邮箱密码\"，\"约会主题\"，\"约会地点\"，\"约会开始时间(可为空)\"，\"约会结束时间(可为空)\"，\"约会类型(可选类型Daily,Weekly,Monthly,Yearly)\"，\"重复约会结束日期\"，\"约会正文\"，以逗号分隔并以换行符结束！");
			System.out.println("例如：");
			System.out.println("  每天事件 usr01@mail1.com,pwd,Daily Recurring Appointments,Tianjin,2018-04-10 10:00:00,2018-04-10 11:00:00,Daily,2018-04-25,Daily appointment body.");
			System.out.println("  每周事件 usr01@mail1.com,pwd,Weekly Recurring Appointments,Tianjin,2018-04-09 13:00:00,2018-04-09 13:50:00,Weekly,2018-05-10,Weekly appointment body.");
			System.out.println("  每月事件 usr01@mail1.com,pwd,Monthly Recurring Appointments,Tianjin,,,Monthly,2018-08-10,Monthly appointment body.");
			System.out.println("  每年事件 usr01@mail1.com,pwd,Yearly Recurring Appointments,Tianjin,,,Yearly,2020-05-01,Yearly appointment body.");
			try {
				input = bfReader.readLine();
				if (input.equals("") || input.equals(null)) {
					throw new IOException("输入无效！");
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(",");
			if (parameters.length == 9) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String subject = parameters[2].trim();
				String location = parameters[3].trim();
				String startsTime = parameters[4].trim();
				String endsTime = parameters[5].trim();
				String recurrenceType = parameters[6].trim();
				String recurrenceEndsDate = parameters[7].trim();
				String body = parameters[8].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.createRecurringAppointment(userName, password, domain, subject, location, startsTime, endsTime, recurrenceType, recurrenceEndsDate, body);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 3: {
			System.out.println("已选择日历操作：[3-创建单一会议]");
			System.out.println("请输入\"会议发起者邮件地址\"，\"会议发起者邮箱密码\"，\"会议必须邀请人\"，\"会议可选邀请人\"，\"会议主题\"，\"会议地点\"，\"会议开始时间(可为空)\"，\"会议结束时间(可为空)\"，\"会议正文\"，以逗号分隔并以换行符结束！");
			System.out.println("例如：usr01@mail1.com,pwd,usr02@mail1.com;qa01@ex-cicc.com,qa02@ex-cicc.com,Single Meeting,Tianjin,2018-04-10 10:00:00,2018-04-10 11:00:00,Meeting Body.");
			try {
				input = bfReader.readLine();
				if (input.equals("") || input.equals(null)) {
					throw new IOException("输入无效！");
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(",");
			if (parameters.length == 9) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String required = parameters[2].trim();
				String optional = parameters[3].trim();
				String subject = parameters[4].trim();
				String location = parameters[5].trim();
				String startsTime = parameters[6].trim();
				String endsTime = parameters[7].trim();
				String body = parameters[8].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.createSingleMeeting(userName, password, domain, required, optional, subject, location, startsTime, endsTime, body);
				} else {
					System.out.println("输入无效！");
				}
			} else {
				System.out.println("输入无效！");
			}
			break;
		}
		case 4: {
			System.out.println("已选择日历操作：[4-创建重复会议]");
			System.out.println("请输入\"会议发起者邮件地址\"，\"会议发起者邮箱密码\"，\"会议必须邀请人\"，\"会议可选邀请人\"，\"会议主题\"，\"会议地点\"，\"会议开始时间(可为空)\"，\"会议结束时间(可为空)\"，\"会议类型(可选类型Daily,Weekly,Monthly,Yearly)\"，\"重复会议结束日期\"，\"会议正文\"，以逗号分隔并以换行符结束！");
			System.out.println("例如：");
			System.out.println("  每天事件 usr01@mail1.com,pwd,usr02@mail1.com;usr03@mail1.com,usr04@mail1.com,Daily Recurring Meeting,Tianjin,2018-04-10 10:00:00,2018-04-10 11:00:00,Daily,2018-04-25,Daily meeting body.");
			System.out.println("  每周事件 usr01@mail1.com,pwd,usr02@mail1.com;usr03@mail1.com,usr04@mail1.com,Weekly Recurring Meeting,Tianjin,2018-04-09 13:00:00,2018-04-09 13:50:00,Weekly,2018-05-10,Weekly meeting body.");
			System.out.println("  每月事件 usr01@mail1.com,pwd,usr02@mail1.com;usr03@mail1.com,usr04@mail1.com,Monthly Recurring Meeting,Tianjin,,,Monthly,2018-08-10,Monthly meeting body.");
			System.out.println("  每年事件 usr01@mail1.com,pwd,usr02@mail1.com;usr03@mail1.com,usr04@mail1.com,Yearly Recurring Meeting,Tianjin,,,Yearly,2020-05-01,Yearly meeting body.");
			try {
				input = bfReader.readLine();
				if (input.equals("") || input.equals(null)) {
					throw new IOException("输入无效！");
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			String[] parameters = input.split(",");
			if (parameters.length == 11) {
				String emailAddress = parameters[0].trim();
				String password = parameters[1].trim();
				String required = parameters[2].trim();
				String optional = parameters[3].trim();
				String subject = parameters[4].trim();
				String location = parameters[5].trim();
				String startsTime = parameters[6].trim();
				String endsTime = parameters[7].trim();
				String recurrenceType = parameters[8].trim();
				String recurrenceEndsDate = parameters[9].trim();
				String body = parameters[10].trim();
				boolean isEmailAddressFormatCorrect = Format.emailFormatCheck(emailAddress);
				if (isEmailAddressFormatCorrect) {
					splitEmailAddress(emailAddress);
					this.createRecurringMeeting(userName, password, domain, required, optional, subject, location, startsTime, endsTime, recurrenceType, recurrenceEndsDate, body);
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
	
	private Date getRecurrenceEndsDate(String recurrenceEndsDate) {
		try {
			// 约会循环类型
			SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
			Date dateRecurrenceEnds = formatter.parse(recurrenceEndsDate);
			return dateRecurrenceEnds;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
	private HashMap<String, Date> getStartsAndEndsTime(String startsTime, String endsTime) {
		HashMap<String, Date> times = new HashMap<String, Date>();
		try {
			SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.CHINA);
			Date timeStarts, timeEnds;
			if ((startsTime.equals("") || startsTime.equals(null)) || (endsTime.equals("") || endsTime.equals(null))) {
				TimeZone.setDefault(TimeZone.getTimeZone("UTC"));
				System.out.println("采用默认时间，事件持续30min...");
				// 使用默认时间
				timeStarts = new Date(); // 约会开始时间
				startsTime = formatter.format(timeStarts);
				timeEnds = new Date(timeStarts.getTime() + 60000*30); // 默认约会Duration为30min
				endsTime = formatter.format(timeEnds);
				System.out.println("默认事件起止时间为: " + startsTime + " / " + endsTime);
			} else {
				TimeZone.setDefault(TimeZone.getTimeZone("GMT"));
				System.out.println("自定义事件起止时间为: " + startsTime + " / " + endsTime);
			}
			timeStarts = formatter.parse(startsTime);
			times.put("startsTime", timeStarts);
			timeEnds = formatter.parse(endsTime);
			times.put("endsTime", timeEnds);
			System.out.println("[Debug] 事件起止时间为: " + timeStarts + " / " + timeEnds);
			return times;
		} catch (Exception e) {
			e.printStackTrace();
		} 
		return null;
	}
}

