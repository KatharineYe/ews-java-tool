package com.ewslib.tools;

import java.util.Date;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import java.text.SimpleDateFormat;


public final class Format {
	/**
	 * 邮件地址格式检查
	 * @param emailAddress
	 * @return
	 */
	public static boolean emailFormatCheck(String emailAddress) {
		String reg = "[a-zA-Z0-9_]+@[a-zA-Z0-9\\-]+(\\.[a-zA-Z]+)+";
		boolean isEmailAddress = emailAddress.matches(reg);
	    if(!isEmailAddress) {
	    	System.out.print("错误：\"" + emailAddress + "\"格式不正确! ");
	    }
	    return isEmailAddress;  
	}
	
	/**
	 * 数字格式检查
	 * @param numStr
	 * @return
	 */
	public static boolean numberCheck(String numStr) {
		boolean isNum = numStr.matches("[0-9]+"); 
		if(!isNum) {
		    System.out.print("错误：\"" + numStr + "\"格式不正确! ");
	    }
		return isNum;
	}
}  

 


