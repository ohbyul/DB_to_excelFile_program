package com.xii.AIHUB;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DBcon {
	String driver = "org.mariadb.jdbc.Driver";
	Connection conn = null;		// DB 접속 객체선언
	PreparedStatement pstmt = null;		// sql 실행할 객체 선언
	ResultSet rs = null;		// sql 실행결과를 담을 객체 선언

	public DBcon() {
		try {
			Class.forName(driver);
			conn = DriverManager.getConnection("jdbc:mariadb://아아피:포트/디비", "아디", "비번");

			if (conn != null) {
				System.out.println("DB 접속 성공 " + conn);
				
				selectList();
			}

		} catch (ClassNotFoundException e) {
			System.out.println("드라이버 로드 실패");
		} catch (SQLException e) {
			System.out.println("DB 접속 실패");
			e.printStackTrace();
		}
	}

	
	public static void main(String[] args) {
		DBcon dbcon = new DBcon();
	}
	
	
	public void selectList() throws SQLException {
		
		String strat_date_sql = "2021/08/01";
		
		SimpleDateFormat format_sql = new SimpleDateFormat("yyyy/MM/dd");
		Date todaytDate_sql = new Date();
		String today_sql = format_sql.format(todaytDate_sql);
		
		String sql_worker = "SELECT USER_ID\r\n"
				+ "	  ,ACCOUNT\r\n"
				+ "	  ,WORK_DATE\r\n"
				+ "	  ,FIRST_WORK_DATE\r\n"
				+ "	  ,IMAGE_ASSIGN_CNT\r\n"
				+ "	  ,(SELECT SUM(TB2.IMAGE_ASSIGN_CNT) FROM TB_AIHUB_PROGRESS AS TB2 WHERE TB2.USER_ID = T1000.USER_ID AND TB2.WORK_DATE <= T1000.WORK_DATE GROUP BY T1000.USER_ID, T1000.WORK_DATE) AS SUM_ASSIGN_CNT\r\n"
				+ "	  ,IMG_CNT\r\n"
				+ "	  ,IMG_INSPECT_1_CNT\r\n"
				+ "	  ,IMG_INSPECT_2_CNT\r\n"
				+ "	  ,IMG_REJECT_CNT\r\n"
				+ "	  ,WORK_START_TIME\r\n"
				+ "	  ,WORK_END_TIME\r\n"
				+ "   FROM (  --  T1000\r\n"
				+ "         SELECT \r\n"
				+ "                T200.USER_ID \r\n"
				+ "               ,T200.ACCOUNT\r\n"
				+ "               ,T200.yyyymmdd  AS WORK_DATE\r\n"
				+ "               ,(SELECT MIN(DATE_FORMAT(CREATED_DATE,'%Y/%m/%d')) FROM META WHERE USER_ID = T200.USER_ID ) AS FIRST_WORK_DATE\r\n"
				+ "               ,IMAGE_ASSIGN_CNT, IMG_CNT,IMG_INSPECT_1_CNT, IMG_INSPECT_2_CNT, IMG_REJECT_CNT, WORK_START_TIME, WORK_END_TIME, WEEK(T200.yyyymmdd ) AS WEEK_CNT\r\n"
				+ "           FROM \r\n"
				+ "              (    -- T100\r\n"
				+ "                 SELECT USER_ID, USER_NAME, FIRST_WORK_DATE, WORK_DATE, IMAGE_ASSIGN_CNT, IMG_CNT,IMG_INSPECT_1_CNT, IMG_INSPECT_2_CNT,IMG_REJECT_CNT, WORK_START_TIME, WORK_END_TIME\r\n"
				+ "                   FROM TB_AIHUB_PROGRESS\r\n"
				+ "                  WHERE WORK_DATE BETWEEN '"+strat_date_sql+"' AND '"+today_sql+"'\r\n"
				+ "               ) T100  \r\n"
				+ "          RIGHT JOIN  \r\n"
				+ "               (    -- T200\r\n"
				+ "                SELECT * \r\n"
				+ "              FROM (    -- T30 \r\n"
				+ "                     SELECT * \r\n"
				+ "                       FROM (   -- T10\r\n"
				+ "                             SELECT n, yyyymmdd\r\n"
				+ "                               FROM ( -- T1\r\n"
				+ "                                     SELECT @N := @N +1 AS n ,  \r\n"
				+ "                                            DATE_FORMAT( DATE_ADD( '2021-07-01' , interval @N -1 day),'%Y/%m/%d') as yyyymmdd  \r\n"
				+ "                                       FROM (aihub.`DATA`  ), (SELECT @N:=0 FROM dual ) a  \r\n"
				+ "                                      LIMIT 500  \r\n"
				+ "                               ) T1\r\n"
				+ "                              WHERE yyyymmdd <= '2021/12/31' \r\n"
				+ "                             ) T10\r\n"
				+ "                       WHERE yyyymmdd BETWEEN '"+strat_date_sql+"' AND '"+today_sql+"' \r\n"
				+ "                  ) T30                   \r\n"
				+ "                   CROSS JOIN ( SELECT USER_ID, ACCOUNT FROM USER WHERE LEVEL_CD = 2 ) T20\r\n"
				+ "                ) T200 \r\n"
				+ "             ON T100.WORK_DATE = T200.yyyymmdd \r\n"
				+ "            AND T100.USER_ID = T200.USER_ID\r\n"
				+ "        ) T1000\r\n"
				+ "   ORDER BY USER_ID, WEEK_CNT, WORK_DATE";
		
		
		String sql_ins = "SELECT USER_ID\r\n"
				+ "	  ,ACCOUNT\r\n"
				+ "	  ,CASE WHEN LEVEL_CD = 3 THEN '1차 검수자'\r\n"
				+ "	  		WHEN LEVEL_CD = 4 THEN '2차 검수자'	  		\r\n"
				+ "	    END AS LEVEL_CD\r\n"
				+ "	  ,WORK_DATE\r\n"
				+ "	  ,FIRST_WORK_DATE\r\n"
				+ "	  ,ASSIGN_CNT\r\n"
				+ "	  ,(SELECT SUM(TB2.ASSIGN_CNT) FROM TB_AIHUB_COMFIRM AS TB2 WHERE TB2.INSPECTOR_ID = T1000.USER_ID AND TB2.CREATED_DATE <= T1000.WORK_DATE GROUP BY T1000.USER_ID, T1000.WORK_DATE) AS SUM_ASSIGN_CNT	-- 누적 할당량\r\n"
				+ "	  ,INS_COMPLETE_1_CNT\r\n"
				+ "	  ,INS_COMPLETE_2_CNT\r\n"
				+ "	  ,INS_REJECT_CNT\r\n"
				+ "	  ,(SELECT SUM(TB2.INS_REJECT_CNT) FROM TB_AIHUB_COMFIRM AS TB2 WHERE TB2.INSPECTOR_ID = T1000.USER_ID AND TB2.CREATED_DATE <= T1000.WORK_DATE GROUP BY T1000.USER_ID, T1000.WORK_DATE) AS SUM_REJECT_CNT	-- 누적 반려량 \r\n"
				+ "	  ,WORK_START_TIME\r\n"
				+ "	  ,WORK_END_TIME\r\n"
				+ "   FROM (  --  T1000\r\n"
				+ "         SELECT \r\n"
				+ "                T200.USER_ID \r\n"
				+ "               ,T200.ACCOUNT\r\n"
				+ "               ,T200.LEVEL_CD\r\n"
				+ "               ,T200.yyyymmdd  AS WORK_DATE\r\n"
				+ "               ,CASE WHEN T200.LEVEL_CD = 3 THEN (SELECT MIN(DATE_FORMAT(DA.CONFIRMED_DATE, '%Y/%m/%d')) FROM DATA DA WHERE DA.CONFIRM_ID = T200.USER_ID  )\r\n"
				+ "					 WHEN T200.LEVEL_CD = 4 THEN (SELECT MIN(DATE_FORMAT(DA.CONFIRMED_DATE2, '%Y/%m/%d')) FROM DATA DA WHERE DA.CONFIRM_ID2 = T200.USER_ID  )\r\n"
				+ "	            END AS FIRST_WORK_DATE 		-- 근무 시작일 / 출퇴근 log 시간은 REJECTED_DATE에 영향을 받지 않습니다. \r\n"
				+ "               ,ASSIGN_CNT, INS_COMPLETE_1_CNT, INS_COMPLETE_2_CNT,INS_REJECT_CNT, WORK_START_TIME, WORK_END_TIME, WEEK(T200.yyyymmdd ) AS WEEK_CNT\r\n"
				+ "           FROM \r\n"
				+ "              (    -- T100\r\n"
				+ "                 SELECT INSPECTOR_ID, ACCOUNT, LEVEL_CD, FIRST_WORK_DATE, CREATED_DATE, ASSIGN_CNT, INS_COMPLETE_1_CNT, INS_COMPLETE_2_CNT ,INS_REJECT_CNT, WORK_START_TIME, WORK_END_TIME\r\n"
				+ "                   FROM TB_AIHUB_COMFIRM\r\n"
				+ "                  WHERE CREATED_DATE BETWEEN '"+strat_date_sql+"' AND '"+today_sql+"'\r\n"
				+ "               ) T100  \r\n"
				+ "          RIGHT JOIN  \r\n"
				+ "               (    -- T200\r\n"
				+ "                SELECT * \r\n"
				+ "              FROM (    -- T30 \r\n"
				+ "                     SELECT * \r\n"
				+ "                       FROM (   -- T10\r\n"
				+ "                             SELECT n, yyyymmdd\r\n"
				+ "                               FROM ( -- T1\r\n"
				+ "                                     SELECT @N := @N +1 AS n ,  \r\n"
				+ "                                            DATE_FORMAT( DATE_ADD( '2021-07-01' , interval @N -1 day),'%Y/%m/%d') as yyyymmdd  \r\n"
				+ "                                       FROM (aihub.`DATA`  ), (SELECT @N:=0 FROM dual ) a  \r\n"
				+ "                                      LIMIT 500  \r\n"
				+ "                               ) T1\r\n"
				+ "                              WHERE yyyymmdd <= '2021/12/31' \r\n"
				+ "                             ) T10\r\n"
				+ "                       WHERE yyyymmdd BETWEEN '"+strat_date_sql+"' AND '"+today_sql+"' \r\n"
				+ "                  ) T30                   \r\n"
				+ "                   CROSS JOIN ( SELECT USER_ID, ACCOUNT,LEVEL_CD  FROM USER WHERE LEVEL_CD IN (3, 4) ) T20\r\n"
				+ "                ) T200 \r\n"
				+ "             ON T100.CREATED_DATE = T200.yyyymmdd \r\n"
				+ "            AND T100.INSPECTOR_ID = T200.USER_ID\r\n"
				+ "        ) T1000\r\n"
				+ "   ORDER BY LEVEL_CD, USER_ID, WEEK_CNT, WORK_DATE";

		
		String sql_category = "	SELECT CATEGORY\r\n"
				+ "	     ,SUM(CASE WHEN CONFIRM_STATUS = 3 THEN INSPECT_CNT\r\n"
				+ "	           ELSE 0 END) AS INSPECT_1_CNT\r\n"
				+ "	     ,SUM(CASE WHEN CONFIRM_STATUS = 4 THEN INSPECT_CNT\r\n"
				+ "	           ELSE 0 END) AS INSPECT_2_CNT\r\n"
				+ "	FROM ( \r\n"
				+ "		   SELECT LA.NAME AS CATEGORY\r\n"
				+ "		           ,DA.CONFIRM_STATUS\r\n"
				+ "		           ,CASE WHEN DA.CONFIRM_STATUS = 3 THEN COUNT(LA.LABEL_ID) \r\n"
				+ "		                 ELSE COUNT(LA.LABEL_ID) END AS INSPECT_CNT \r\n"
				+ "		    FROM LABEL LA\r\n"
				+ "		    JOIN META ME \r\n"
				+ "		      ON LA.LABEL_ID = ME.LABEL_ID\r\n"
				+ "		    JOIN DATA DA \r\n"
				+ "		      ON ME.DATA_ID = DA.DATA_ID\r\n"
				+ "		   WHERE DA.STATUS = 1\r\n"
				+ "		     AND DA.CONFIRM_STATUS IN ( 3, 4 ) -- 1차 검수 완료\r\n"
				+ "		     AND CASE WHEN DA.CONFIRM_STATUS = 3 THEN DA.CONFIRMED_DATE IS NOT NULL \r\n"
				+ "		              ELSE DA.CONFIRMED_DATE2 IS NOT NULL END\r\n"
				+ "		 GROUP BY LA.NAME, DA.CONFIRM_STATUS\r\n"
				+ "	 ) AAA\r\n"
				+ "	 GROUP BY CATEGORY";
		
		Statement stmt = conn.createStatement();
		rs =stmt.executeQuery(sql_worker);		//SQL 수행 후 객체 생성
		
		ArrayList<String[]> workerList = new ArrayList<String[]>();
		
		while (rs.next()) {
		  String[] arrStr = {rs.getString("USER_ID"),rs.getString("ACCOUNT"),rs.getString("WORK_DATE"),rs.getString("FIRST_WORK_DATE"),rs.getString("IMAGE_ASSIGN_CNT"),rs.getString("SUM_ASSIGN_CNT"),rs.getString("IMG_CNT"),rs.getString("IMG_INSPECT_1_CNT"),rs.getString("IMG_INSPECT_2_CNT"),rs.getString("IMG_REJECT_CNT"),rs.getString("WORK_START_TIME"),rs.getString("WORK_END_TIME")};
		  workerList.add(arrStr);
		}
		
		for (int i = 0; i < workerList.size(); i++) {
			  System.out.println("USER_ID : " + workerList.get(i)[0]);
			  System.out.println("ACCOUNT : " + workerList.get(i)[1]);
			  System.out.println("WORK_DATE : " + workerList.get(i)[2]);
			  System.out.println("FIRST_WORK_DATE : " + workerList.get(i)[3]);
			  System.out.println("IMAGE_ASSIGN_CNT : " + workerList.get(i)[4]);
			  System.out.println("SUM_ASSIGN_CNT : " + workerList.get(i)[5]);
			  System.out.println("IMG_CNT : " + workerList.get(i)[6]);
			  System.out.println("IMG_INSPECT_1_CNT : " + workerList.get(i)[7]);
			  System.out.println("IMG_INSPECT_2_CNT : " + workerList.get(i)[8]);
			  System.out.println("IMG_REJECT_CNT : " + workerList.get(i)[9]);
			  System.out.println("WORK_START_TIME : " + workerList.get(i)[10]);
			  System.out.println("WORK_END_TIME : " + workerList.get(i)[11]);
		}
		
		rs =stmt.executeQuery(sql_ins);		//SQL 수행 후 객체 생성
		
		ArrayList<String[]> insList = new ArrayList<String[]>();
		
		while (rs.next()) {
		  String[] arrStr_ins = {rs.getString("USER_ID"),rs.getString("ACCOUNT"),rs.getString("LEVEL_CD"),rs.getString("WORK_DATE"),rs.getString("FIRST_WORK_DATE"),rs.getString("ASSIGN_CNT"),rs.getString("SUM_ASSIGN_CNT"),rs.getString("INS_COMPLETE_1_CNT"),rs.getString("INS_COMPLETE_2_CNT"),rs.getString("INS_REJECT_CNT"),rs.getString("SUM_REJECT_CNT"),rs.getString("WORK_START_TIME"),rs.getString("WORK_END_TIME")};
		  insList.add(arrStr_ins);
		}
		
		for (int i = 0; i < insList.size(); i++) {
			  System.out.println("USER_ID : " + insList.get(i)[0]);
			  System.out.println("ACCOUNT : " + insList.get(i)[1]);
			  System.out.println("LEVEL_CD : " + insList.get(i)[2]);
			  System.out.println("WORK_DATE : " + insList.get(i)[3]);
			  System.out.println("FIRST_WORK_DATE : " + insList.get(i)[4]);
			  System.out.println("ASSIGN_CNT : " + insList.get(i)[5]);
			  System.out.println("SUM_ASSIGN_CNT : " + insList.get(i)[6]);
			  System.out.println("INS_COMPLETE_1_CNT : " + insList.get(i)[7]);
			  System.out.println("INS_COMPLETE_2_CNT : " + insList.get(i)[8]);
			  System.out.println("IMG_REJECT_CNT : " + insList.get(i)[9]);
			  System.out.println("SUM_REJECT_CNT : " + insList.get(i)[10]);
			  System.out.println("WORK_START_TIME : " + insList.get(i)[11]);
			  System.out.println("WORK_END_TIME : " + insList.get(i)[12]);
		}
		
		
		rs =stmt.executeQuery(sql_category);		//SQL 수행 후 객체 생성
		
		ArrayList<String[]> categoryList = new ArrayList<String[]>();
		
		while (rs.next()) {
		  String[] arrStr_ca = {rs.getString("CATEGORY"),rs.getString("INSPECT_1_CNT"),rs.getString("INSPECT_2_CNT")};
		  categoryList.add(arrStr_ca);
		}

		for (int i = 0; i < categoryList.size(); i++) {
			  System.out.println("CATEGORY : " + categoryList.get(i)[0]);
			  System.out.println("INSPECT_1_CNT : " + categoryList.get(i)[1]);
			  System.out.println("INSPECT_2_CNT : " + categoryList.get(i)[2]);
		}
		
		writeExcel(workerList,insList,categoryList);
		
		rs.close();
		stmt.close();
		conn.close();
	}
	
	
	public void writeExcel(ArrayList<String[]> workerList,ArrayList<String[]> insList,ArrayList<String[]> categoryList) {

		
		SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
		Date todaytDate = new Date();
		String today = format.format(todaytDate);
		
		//.xlsx 확장자 지원
		XSSFWorkbook xssfWb = null; 
		XSSFSheet xssfSheet = null; 	//작업자 워크시트
		XSSFSheet xssfSheet_ins = null; 	//작업자 워크시트
		XSSFSheet xssfSheet_ca = null; 	//작업자 워크시트
		XSSFRow xssfRow = null; 
		XSSFRow xssfRow_ins = null; 
		XSSFRow xssfRow_ca = null; 
		XSSFCell xssfCell = null;
		XSSFCell xssfCell_ins = null;
		XSSFCell xssfCell_ca = null;
		
		try {
			int rowNo = 0; // 행 갯수 
			// 워크북 생성
			xssfWb = new XSSFWorkbook();
			xssfSheet = xssfWb.createSheet("AI Hub-51 Worker"); // 워크시트 이름
			
			
			xssfSheet.setColumnWidth(0, (xssfSheet.getColumnWidth(0))+(short)2048); // 0번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(1, (xssfSheet.getColumnWidth(1))+(short)2048); // 1번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(2, (xssfSheet.getColumnWidth(2))+(short)2048); // 2번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(3, (xssfSheet.getColumnWidth(3))+(short)2048); // 3번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(4, (xssfSheet.getColumnWidth(4))+(short)2048); // 4번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(5, (xssfSheet.getColumnWidth(5))+(short)2048); // 5번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(6, (xssfSheet.getColumnWidth(6))+(short)2048); // 6번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(7, (xssfSheet.getColumnWidth(7))+(short)2048); // 7번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(8, (xssfSheet.getColumnWidth(8))+(short)2048); // 8번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(9, (xssfSheet.getColumnWidth(9))+(short)2048); // 9번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(10, (xssfSheet.getColumnWidth(10))+(short)4096); // 10번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(11, (xssfSheet.getColumnWidth(11))+(short)4096); // 11번째 컬럼 넓이 조절
			
			//헤더용 폰트 스타일
			XSSFFont font = xssfWb.createFont();
			font.setFontName(HSSFFont.FONT_ARIAL); //폰트스타일
			font.setFontHeightInPoints((short)20); //폰트크기
			font.setBold(true); //Bold 유무
			
			//테이블 타이틀 스타일
			CellStyle cellStyle_Title = xssfWb.createCellStyle();
			cellStyle_Title.setBorderTop(BorderStyle.THIN); //테두리 위쪽
			cellStyle_Title.setBorderBottom(BorderStyle.THIN); //테두리 아래쪽
			cellStyle_Title.setBorderLeft(BorderStyle.THIN); //테두리 왼쪽
			cellStyle_Title.setBorderRight(BorderStyle.THIN); //테두리 오른쪽
			cellStyle_Title.setFont(font); // cellStle에 font를 적용
			cellStyle_Title.setAlignment(HorizontalAlignment.CENTER); // 정렬
			
			//셀병합
			xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11)); //첫행, 마지막행, 첫열, 마지막열( 0번째 행의 0~11번째 컬럼을 병합한다)
			
			//타이틀 생성
			xssfRow = xssfSheet.createRow(rowNo++); //행 객체 추가
			xssfCell = xssfRow.createCell((short) 0); // 추가한 행에 셀 객체 추가
			xssfCell.setCellStyle(cellStyle_Title); // 셀에 스타일 지정
			xssfCell.setCellValue("AI HUB - 51 Worker 능률 " + today ); // 데이터 입력
			
			//xssfRow = xssfSheet.createRow(rowNo++);  // 빈행 추가
			
			//컬럼 용 폰트 스타일
			XSSFFont font2 = xssfWb.createFont();
			font2.setFontName(HSSFFont.FONT_ARIAL); //폰트스타일
			font2.setFontHeightInPoints((short)12); //폰트크기
			font2.setBold(true); //Bold 유무
			
			CellStyle cellStyle_Body = xssfWb.createCellStyle(); 
			cellStyle_Body.setBorderTop(BorderStyle.THIN); //테두리 위쪽
			cellStyle_Body.setBorderBottom(BorderStyle.THIN); //테두리 아래쪽
			cellStyle_Body.setBorderLeft(BorderStyle.THIN); //테두리 왼쪽
			cellStyle_Body.setBorderRight(BorderStyle.THIN); //테두리 오른쪽
			cellStyle_Body.setFont(font2); // cellStle에 font를 적용
		
			
			
			//content 용 폰트 스타일
			XSSFFont font3 = xssfWb.createFont();
			font3.setFontName(HSSFFont.FONT_ARIAL); //폰트스타일
			font3.setFontHeightInPoints((short)9); //폰트크기
			
			CellStyle cellStyle_content = xssfWb.createCellStyle(); 
			cellStyle_content.setBorderTop(BorderStyle.THIN); //테두리 위쪽
			cellStyle_content.setBorderBottom(BorderStyle.THIN); //테두리 아래쪽
			cellStyle_content.setBorderLeft(BorderStyle.THIN); //테두리 왼쪽
			cellStyle_content.setBorderRight(BorderStyle.THIN); //테두리 오른쪽
			cellStyle_content.setFont(font3); // cellStle에 font를 적용
			
			
			
			//헤더 생성
			xssfRow = xssfSheet.createRow(rowNo++); //헤더 01

			
			xssfCell = xssfRow.createCell((short) 0);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("USER_ID");
			
			xssfCell = xssfRow.createCell((short) 1);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("ACCOUNT");
			
			xssfCell = xssfRow.createCell((short) 2);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("작업일");

			xssfCell = xssfRow.createCell((short) 3);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("근무 시작일");

			xssfCell = xssfRow.createCell((short) 4);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("할당량(IMG)");
			
			xssfCell = xssfRow.createCell((short) 5);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("누적 할당량");
			
			xssfCell = xssfRow.createCell((short) 6);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("작업량(IMG)");
			
			xssfCell = xssfRow.createCell((short) 7);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("1차검수완료량");
			
			xssfCell = xssfRow.createCell((short) 8);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("2차검수완료량");
			
			xssfCell = xssfRow.createCell((short) 9);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("반려량");
			
			xssfCell = xssfRow.createCell((short) 10);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("START_TIME");
			
			xssfCell = xssfRow.createCell((short) 11);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("END_TIME");		
		
			for(int i = 0 ;i<workerList.size();i++) {
				xssfRow = xssfSheet.createRow(rowNo++); //헤더 02
				xssfCell = xssfRow.createCell((short) 0);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[0]);
				
				xssfCell = xssfRow.createCell((short) 1);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[1]);
				
				xssfCell = xssfRow.createCell((short) 2);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[2]);

				xssfCell = xssfRow.createCell((short) 3);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[3]);

				xssfCell = xssfRow.createCell((short) 4);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[4]);

				xssfCell = xssfRow.createCell((short) 5);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[5]);

				xssfCell = xssfRow.createCell((short) 6);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[6]);

				xssfCell = xssfRow.createCell((short) 7);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[7]);

				xssfCell = xssfRow.createCell((short) 8);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[8]);

				xssfCell = xssfRow.createCell((short) 9);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[9]);

				xssfCell = xssfRow.createCell((short) 10);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[10]);
				
				xssfCell = xssfRow.createCell((short) 11);
				xssfCell.setCellStyle(cellStyle_content);
				xssfCell.setCellValue(workerList.get(i)[11]);
			}
/*		---------------------------워크 시트 2 검수자 -----------------------------------------------------------------------		*/			
			xssfSheet_ins = xssfWb.createSheet("AI Hub-51 검수자"); // 워크시트 생성
			int rowNo_ins = 0; // 행 갯수 
			
			xssfSheet_ins.setColumnWidth(0, (xssfSheet_ins.getColumnWidth(0))+(short)2048); // 0번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(1, (xssfSheet_ins.getColumnWidth(1))+(short)2048); // 1번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(2, (xssfSheet_ins.getColumnWidth(2))+(short)2048); // 2번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(3, (xssfSheet_ins.getColumnWidth(3))+(short)2048); // 3번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(4, (xssfSheet_ins.getColumnWidth(4))+(short)2048); // 4번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(5, (xssfSheet_ins.getColumnWidth(5))+(short)2048); // 5번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(6, (xssfSheet_ins.getColumnWidth(6))+(short)2048); // 6번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(7, (xssfSheet_ins.getColumnWidth(7))+(short)2048); // 7번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(8, (xssfSheet_ins.getColumnWidth(8))+(short)2048); // 8번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(9, (xssfSheet_ins.getColumnWidth(9))+(short)2048); // 9번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(10, (xssfSheet_ins.getColumnWidth(10))+(short)2048); // 10번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(11, (xssfSheet_ins.getColumnWidth(11))+(short)4096); // 11번째 컬럼 넓이 조절
			xssfSheet_ins.setColumnWidth(12, (xssfSheet_ins.getColumnWidth(12))+(short)4096); // 12번째 컬럼 넓이 조절
			

			//셀병합
			xssfSheet_ins.addMergedRegion(new CellRangeAddress(0, 0, 0, 12)); //첫행, 마지막행, 첫열, 마지막열( 0번째 행의 0~12번째 컬럼을 병합한다)
			
			//타이틀 생성
			xssfRow_ins = xssfSheet_ins.createRow(rowNo_ins++); //행 객체 추가
			xssfCell_ins = xssfRow_ins.createCell((short) 0); // 추가한 행에 셀 객체 추가
			xssfCell_ins.setCellStyle(cellStyle_Title); // 셀에 스타일 지정
			xssfCell_ins.setCellValue("AI HUB - 51 검수자 능률 " + today ); // 데이터 입력
			
			
			//헤더 생성
			xssfRow_ins = xssfSheet_ins.createRow(rowNo_ins++); //헤더 01

			
			xssfCell_ins = xssfRow_ins.createCell((short) 0);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("USER_ID");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 1);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("ACCOUNT");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 2);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("LEVEL_CD");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 3);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("작업일");

			xssfCell_ins = xssfRow_ins.createCell((short) 4);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("근무 시작일");

			xssfCell_ins = xssfRow_ins.createCell((short) 5);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("할당량(IMG)");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 6);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("누적 할당량");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 7);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("1차검수완료량");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 8);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("2차검수완료량");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 9);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("반려량");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 10);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("누적반려량");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 11);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("START_TIME");
			
			xssfCell_ins = xssfRow_ins.createCell((short) 12);
			xssfCell_ins.setCellStyle(cellStyle_Body);
			xssfCell_ins.setCellValue("END_TIME");		
		
			for(int i = 0 ;i<insList.size();i++) {
				xssfRow_ins = xssfSheet_ins.createRow(rowNo_ins++); //헤더 02
				xssfCell_ins = xssfRow_ins.createCell((short) 0);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[0]);
				
				xssfCell_ins = xssfRow_ins.createCell((short) 1);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[1]);
				
				xssfCell_ins = xssfRow_ins.createCell((short) 2);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[2]);

				xssfCell_ins = xssfRow_ins.createCell((short) 3);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[3]);

				xssfCell_ins = xssfRow_ins.createCell((short) 4);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[4]);

				xssfCell_ins = xssfRow_ins.createCell((short) 5);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[5]);

				xssfCell_ins = xssfRow_ins.createCell((short) 6);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[6]);

				xssfCell_ins = xssfRow_ins.createCell((short) 7);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[7]);

				xssfCell_ins = xssfRow_ins.createCell((short) 8);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[8]);

				xssfCell_ins = xssfRow_ins.createCell((short) 9);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[9]);

				xssfCell_ins = xssfRow_ins.createCell((short) 10);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[10]);
				
				xssfCell_ins = xssfRow_ins.createCell((short) 11);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[11]);
				
				xssfCell_ins = xssfRow_ins.createCell((short) 12);
				xssfCell_ins.setCellStyle(cellStyle_content);
				xssfCell_ins.setCellValue(insList.get(i)[12]);
			}

/*		---------------------------워크 시트 3 카테고리 -----------------------------------------------------------------------		*/			
			xssfSheet_ca = xssfWb.createSheet("AI Hub-51 카테고리"); // 워크시트 생성
			int rowNo_ca = 0; // 행 갯수 
			
			xssfSheet_ca.setColumnWidth(0, (xssfSheet_ca.getColumnWidth(0))+(short)2048); // 0번째 컬럼 넓이 조절
			xssfSheet_ca.setColumnWidth(1, (xssfSheet_ca.getColumnWidth(1))+(short)2048); // 1번째 컬럼 넓이 조절
			xssfSheet_ca.setColumnWidth(2, (xssfSheet_ca.getColumnWidth(2))+(short)2048); // 2번째 컬럼 넓이 조절


			//셀병합
			xssfSheet_ca.addMergedRegion(new CellRangeAddress(0, 0, 0, 2)); //첫행, 마지막행, 첫열, 마지막열( 0번째 행의 0~2번째 컬럼을 병합한다)
			
			//타이틀 생성
			xssfRow_ca = xssfSheet_ca.createRow(rowNo_ca++); //행 객체 추가
			xssfCell_ca = xssfRow_ca.createCell((short) 0); // 추가한 행에 셀 객체 추가
			xssfCell_ca.setCellStyle(cellStyle_Title); // 셀에 스타일 지정
			xssfCell_ca.setCellValue("카테고리별 1차/2차 검수완료"); // 데이터 입력
			
			
			//헤더 생성
			xssfRow_ca = xssfSheet_ca.createRow(rowNo_ca++); //헤더 01
			
			xssfCell_ca = xssfRow_ca.createCell((short) 0);
			xssfCell_ca.setCellStyle(cellStyle_Body);
			xssfCell_ca.setCellValue("CATEGORY");
			
			xssfCell_ca = xssfRow_ca.createCell((short) 1);
			xssfCell_ca.setCellStyle(cellStyle_Body);
			xssfCell_ca.setCellValue("1차 검수완료량");
			
			xssfCell_ca = xssfRow_ca.createCell((short) 2);
			xssfCell_ca.setCellStyle(cellStyle_Body);
			xssfCell_ca.setCellValue("2차 검수완료량");
			
			for(int i = 0 ;i<categoryList.size();i++) {
				xssfRow_ca = xssfSheet_ca.createRow(rowNo_ca++); //헤더 02
				xssfCell_ca = xssfRow_ca.createCell((short) 0);
				xssfCell_ca.setCellStyle(cellStyle_content);
				xssfCell_ca.setCellValue(categoryList.get(i)[0]);
				
				xssfCell_ca = xssfRow_ca.createCell((short) 1);
				xssfCell_ca.setCellStyle(cellStyle_content);
				xssfCell_ca.setCellValue(categoryList.get(i)[1]);
				
				xssfCell_ca = xssfRow_ca.createCell((short) 2);
				xssfCell_ca.setCellStyle(cellStyle_content);
				xssfCell_ca.setCellValue(categoryList.get(i)[2]);

			}
	
			String path = "C:/test_out/";
					
			String localFile = path + "AI_HUB_51[" + today + "]" + ".xlsx";
			
			File file = new File(localFile);
			FileOutputStream fos = null;
			fos = new FileOutputStream(file);
			xssfWb.write(fos);

			if (xssfWb != null)	xssfWb.close();
			
			System.out.println(localFile + " 출력 완료 ");
			
			}
			catch(Exception e){
				e.printStackTrace();
			}

	}
	
	
	
	
	
	
}



