package com.xii.AIHUB;

import java.text.SimpleDateFormat;
import java.util.Date;

public class date_test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		SimpleDateFormat format_sql = new SimpleDateFormat("yyyy/MM/dd");
		SimpleDateFormat format_all = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		
		Date todaytDate_sql = new Date();
		String today_sql = format_sql.format(todaytDate_sql);
		String today_all = format_all.format(todaytDate_sql);
		
		
		
		System.out.println("today " + today_sql);
		System.out.println("todaytDate_sql " + today_all);
		
		
		Date yesterDay_sql = new Date();
		yesterDay_sql = new Date(yesterDay_sql.getTime()+(1000*60*60*24*-1));
		String yesterday_sql = format_sql.format(yesterDay_sql);
		
		
		System.out.println("yesterday" + yesterday_sql);
		
		
	}

}
