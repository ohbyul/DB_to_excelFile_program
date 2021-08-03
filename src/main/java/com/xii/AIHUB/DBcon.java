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
			conn = DriverManager.getConnection("jdbc:mariadb://127.0.0.1:3306/b_test", "id", "pw");

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
		
		String sql = "SELECT *\r\n"
				+ "   FROM (  --  T1000\r\n"
				+ "         SELECT \r\n"
				+ "                T200.USER_ID \r\n"
				+ "               ,T200.ACCOUNT\r\n"
				+ "               ,T200.yyyymmdd  AS WORK_DATE\r\n"
				+ "               ,(SELECT MIN(DATE_FORMAT(CREATED_DATE,'%Y/%m/%d')) FROM META WHERE USER_ID = T200.USER_ID ) AS FIRST_WORK_DATE\r\n"
				+ "               ,IMAGE_ASSIGN_CNT, IMG_CNT,IMG_INSPECT_CNT, BOX_CNT, BOX_INSPECT_CNT, WORK_START_TIME, WORK_END_TIME, WEEK(T200.yyyymmdd ) AS WEEK_CNT\r\n"
				+ "           FROM \r\n"
				+ "              (    -- T100\r\n"
				+ "                 SELECT USER_ID, USER_NAME, FIRST_WORK_DATE, WORK_DATE, IMAGE_ASSIGN_CNT, IMG_CNT,IMG_INSPECT_CNT, BOX_CNT,BOX_INSPECT_CNT, WORK_START_TIME, WORK_END_TIME\r\n"
				+ "                   FROM TB_AIHUB_PROGRESS\r\n"
				+ "                  WHERE WORK_DATE BETWEEN '2021/07/27' AND '2021/08/02'\r\n"
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
				+ "                                            DATE_FORMAT( DATE_ADD( '2021-07-27' , interval @N -1 day),'%Y/%m/%d') as yyyymmdd  \r\n"
				+ "                                       FROM (aihub.`DATA`  ), (SELECT @N:=0 FROM dual ) a  \r\n"
				+ "                                      LIMIT 500  \r\n"
				+ "                               ) T1\r\n"
				+ "                              WHERE yyyymmdd <= '2021/12/31' \r\n"
				+ "                             ) T10\r\n"
				+ "                       WHERE yyyymmdd BETWEEN '2021/07/27' AND '2021/08/02' \r\n"
				+ "                  ) T30                   \r\n"
				+ "                   CROSS JOIN ( SELECT USER_ID, ACCOUNT FROM USER WHERE LEVEL_CD = 2 ) T20\r\n"
				+ "                ) T200 \r\n"
				+ "             ON T100.WORK_DATE = T200.yyyymmdd \r\n"
				+ "            AND T100.USER_ID = T200.USER_ID\r\n"
				+ "      UNION ALL      -- 월간 합계\r\n"
				+ "         SELECT TB.USER_ID, TB.USER_NAME, '사용자별 월간 합계' AS WORK_DATE, TB.FIRST_WORK_DATE,  SUM(TB.IMAGE_ASSIGN_CNT), SUM(TB.IMG_CNT), SUM(TB.IMG_INSPECT_CNT), SUM(TB.BOX_CNT), SUM(TB.BOX_INSPECT_CNT), MIN(TB.WORK_START_TIME), MAX(TB.WORK_END_TIME) , 999\r\n"
				+ "           FROM TB_AIHUB_PROGRESS TB \r\n"
				+ "           JOIN USER US \r\n"
				+ "             ON TB.USER_ID = US.USER_ID\r\n"
				+ "            AND US.LEVEL_CD =2 \r\n"
				+ "          WHERE WORK_DATE BETWEEN '2021/07/27' AND '2021/08/02'\r\n"
				+ "       GROUP BY TB.USER_ID, TB.FIRST_WORK_DATE\r\n"
				+ "      UNION ALL      -- 주간 소계\r\n"
				+ "      SELECT \r\n"
				+ "          TT200.USER_ID \r\n"
				+ "         ,TT200.ACCOUNT AS USER_NAME   -- ACCOUNT 인지..\r\n"
				+ "         ,'주간소계' AS WORK_DATE\r\n"
				+ "         ,FIRST_WORK_DATE\r\n"
				+ "         ,SUM(IMAGE_ASSIGN_CNT), SUM(IMG_CNT),SUM(IMG_INSPECT_CNT), SUM(BOX_CNT), SUM(BOX_INSPECT_CNT), MIN(WORK_START_TIME), MAX(WORK_END_TIME), WEEK(TT200.yyyymmdd)\r\n"
				+ "         FROM \r\n"
				+ "            (    -- TT100\r\n"
				+ "           SELECT USER_ID, USER_NAME, FIRST_WORK_DATE, WORK_DATE, IMAGE_ASSIGN_CNT, IMG_CNT,IMG_INSPECT_CNT, BOX_CNT,BOX_INSPECT_CNT, WORK_START_TIME, WORK_END_TIME\r\n"
				+ "             FROM TB_AIHUB_PROGRESS\r\n"
				+ "            WHERE WORK_DATE BETWEEN '2021/07/27' AND '2021/08/02'\r\n"
				+ "             ) TT100  \r\n"
				+ "        RIGHT JOIN  \r\n"
				+ "             (    -- TT200\r\n"
				+ "          SELECT * \r\n"
				+ "        FROM (    -- TT30 \r\n"
				+ "               SELECT * \r\n"
				+ "                 FROM (   -- TT10\r\n"
				+ "                       SELECT n, yyyymmdd\r\n"
				+ "                         FROM ( -- TT1\r\n"
				+ "                               SELECT @NN := @NN +1 AS n ,  \r\n"
				+ "                                      DATE_FORMAT( DATE_ADD( '2021-07-27' , interval @NN -1 day),'%Y/%m/%d') as yyyymmdd  \r\n"
				+ "                                 FROM (aihub.`DATA`  ), (SELECT @NN:=0 FROM dual ) a  \r\n"
				+ "                                LIMIT 500  \r\n"
				+ "                         ) TT1\r\n"
				+ "                        WHERE yyyymmdd <= '2021/12/31' \r\n"
				+ "                       ) TT10\r\n"
				+ "                 WHERE yyyymmdd BETWEEN '2021/07/27' AND '2021/08/02' \r\n"
				+ "            ) TT30                   \r\n"
				+ "             CROSS JOIN ( SELECT USER_ID, ACCOUNT FROM USER WHERE LEVEL_CD = 2 ) TT20\r\n"
				+ "          ) TT200 \r\n"
				+ "       ON TT100.WORK_DATE = TT200.yyyymmdd \r\n"
				+ "      AND TT100.USER_ID = TT200.USER_ID\r\n"
				+ "      GROUP BY TT200.USER_ID , WEEK(TT200.yyyymmdd)  -- 주간 소계 끝\r\n"
				+ "        ) T1000\r\n"
				+ "      WHERE USER_ID >= 34 \r\n"
				+ "   ORDER BY USER_ID, WEEK_CNT, WORK_DATE";
		
		Statement stmt = conn.createStatement();
		rs =stmt.executeQuery(sql);		//SQL 수행 후 객체 생성
		
		ArrayList<String[]> jobList = new ArrayList<String[]>();
		
		while (rs.next()) {
		  String[] arrStr = {rs.getString("USER_ID"),rs.getString("ACCOUNT"),rs.getString("WORK_DATE"),rs.getString("FIRST_WORK_DATE"),rs.getString("IMAGE_ASSIGN_CNT"),rs.getString("IMG_CNT"),rs.getString("IMG_INSPECT_CNT"),rs.getString("BOX_CNT"),rs.getString("BOX_INSPECT_CNT"),rs.getString("WORK_START_TIME"),rs.getString("WORK_END_TIME")};
		  jobList.add(arrStr);
		}
		
		for (int i = 0; i < jobList.size(); i++) {
			  System.out.println("USER_ID : " + jobList.get(i)[0]);
			  System.out.println("ACCOUNT : " + jobList.get(i)[1]);
			  System.out.println("WORK_DATE : " + jobList.get(i)[2]);
			  System.out.println("FIRST_WORK_DATE : " + jobList.get(i)[3]);
			  System.out.println("IMAGE_ASSIGN_CNT : " + jobList.get(i)[4]);
			  System.out.println("IMG_CNT : " + jobList.get(i)[5]);
			  System.out.println("IMG_INSPECT_CNT : " + jobList.get(i)[6]);
			  System.out.println("BOX_CNT : " + jobList.get(i)[7]);
			  System.out.println("BOX_INSPECT_CNT : " + jobList.get(i)[8]);
			  System.out.println("WORK_START_TIME : " + jobList.get(i)[9]);
			  System.out.println("WORK_END_TIME : " + jobList.get(i)[10]);
		}

		writeExcel(jobList);
		
		rs.close();
		stmt.close();
		conn.close();
	}
	
	
	public void writeExcel(ArrayList<String[]> list) {

		
		SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
		Date todaytDate = new Date();
		String today = format.format(todaytDate);
		
		//.xlsx 확장자 지원
		XSSFWorkbook xssfWb = null; 
		XSSFSheet xssfSheet = null; 
		XSSFRow xssfRow = null; 
		XSSFCell xssfCell = null;
		
		try {
			int rowNo = 0; // 행 갯수 
			// 워크북 생성
			xssfWb = new XSSFWorkbook();
			xssfSheet = xssfWb.createSheet("AI Hub-51 Worker"); // 워크시트 이름
			
			//헤더용 폰트 스타일
			XSSFFont font = xssfWb.createFont();
			font.setFontName(HSSFFont.FONT_ARIAL); //폰트스타일
			font.setFontHeightInPoints((short)14); //폰트크기
			font.setBold(true); //Bold 유무
			
			//테이블 타이틀 스타일
			CellStyle cellStyle_Title = xssfWb.createCellStyle();
			
			xssfSheet.setColumnWidth(0, (xssfSheet.getColumnWidth(0))+(short)2048); // 0번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(1, (xssfSheet.getColumnWidth(1))+(short)2048); // 1번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(2, (xssfSheet.getColumnWidth(2))+(short)4096); // 2번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(3, (xssfSheet.getColumnWidth(3))+(short)4096); // 3번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(4, (xssfSheet.getColumnWidth(4))+(short)2048); // 4번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(5, (xssfSheet.getColumnWidth(5))+(short)2048); // 5번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(6, (xssfSheet.getColumnWidth(6))+(short)2048); // 6번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(7, (xssfSheet.getColumnWidth(7))+(short)2048); // 7번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(8, (xssfSheet.getColumnWidth(8))+(short)2048); // 8번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(9, (xssfSheet.getColumnWidth(9))+(short)4096); // 9번째 컬럼 넓이 조절
			xssfSheet.setColumnWidth(10, (xssfSheet.getColumnWidth(10))+(short)4096); // 10번째 컬럼 넓이 조절
			
			cellStyle_Title.setFont(font); // cellStle에 font를 적용
			cellStyle_Title.setAlignment(HorizontalAlignment.CENTER); // 정렬
			
			//셀병합
			xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10)); //첫행, 마지막행, 첫열, 마지막열( 0번째 행의 0~2번째 컬럼을 병합한다)
			//타이틀 생성
			xssfRow = xssfSheet.createRow(rowNo++); //행 객체 추가
			xssfCell = xssfRow.createCell((short) 0); // 추가한 행에 셀 객체 추가
			xssfCell.setCellStyle(cellStyle_Title); // 셀에 스타일 지정
			xssfCell.setCellValue("AI HUB - 51 Worker 능률 " + today ); // 데이터 입력
			
			//xssfRow = xssfSheet.createRow(rowNo++);  // 빈행 추가
			
			CellStyle cellStyle_Body = xssfWb.createCellStyle(); 
			cellStyle_Body.setAlignment(HorizontalAlignment.LEFT); 
			
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
			xssfCell.setCellValue("WORK_DATE");

			xssfCell = xssfRow.createCell((short) 3);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("FIRST_WORK_DATE");

			xssfCell = xssfRow.createCell((short) 4);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("IMAGE_ASSIGN_CNT");
			
			xssfCell = xssfRow.createCell((short) 5);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("IMG_CNT");
			
			xssfCell = xssfRow.createCell((short) 6);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("IMG_INSPECT_CNT");
			
			xssfCell = xssfRow.createCell((short) 7);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("BOX_CNT");
			
			xssfCell = xssfRow.createCell((short) 8);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("BOX_INSPECT_CNT");
			
			xssfCell = xssfRow.createCell((short) 9);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("WORK_START_TIME");
			
			xssfCell = xssfRow.createCell((short) 10);
			xssfCell.setCellStyle(cellStyle_Body);
			xssfCell.setCellValue("WORK_END_TIME");		
		
			for(int i = 0 ;i<list.size();i++) {
				xssfRow = xssfSheet.createRow(rowNo++); //헤더 02
				xssfCell = xssfRow.createCell((short) 0);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[0]);
				
				xssfCell = xssfRow.createCell((short) 1);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[1]);
				
				xssfCell = xssfRow.createCell((short) 2);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[2]);

				xssfCell = xssfRow.createCell((short) 3);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[3]);

				xssfCell = xssfRow.createCell((short) 4);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[4]);

				xssfCell = xssfRow.createCell((short) 5);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[5]);

				xssfCell = xssfRow.createCell((short) 6);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[6]);

				xssfCell = xssfRow.createCell((short) 7);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[7]);

				xssfCell = xssfRow.createCell((short) 8);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[8]);

				xssfCell = xssfRow.createCell((short) 9);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[9]);

				xssfCell = xssfRow.createCell((short) 10);
				xssfCell.setCellStyle(cellStyle_Body);
				xssfCell.setCellValue(list.get(i)[10]);
			}
			

			//테이블 스타일 설정
			/*
			CellStyle cellStyle_Table_Center = xssfWb.createCellStyle();
			cellStyle_Table_Center.setBorderTop(BorderStyle.THIN); //테두리 위쪽
			cellStyle_Table_Center.setBorderBottom(BorderStyle.THIN); //테두리 아래쪽
			cellStyle_Table_Center.setBorderLeft(BorderStyle.THIN); //테두리 왼쪽
			cellStyle_Table_Center.setBorderRight(BorderStyle.THIN); //테두리 오른쪽
			cellStyle_Table_Center.setAlignment(HorizontalAlignment.CENTER);
			
			xssfRow = xssfSheet.createRow(rowNo++);
			xssfCell = xssfRow.createCell((short) 0);
			xssfCell.setCellStyle(cellStyle_Table_Center);
			xssfCell.setCellValue("테이블 셀1");
			
			xssfCell = xssfRow.createCell((short) 1);
			xssfCell.setCellStyle(cellStyle_Table_Center);
			xssfCell.setCellValue("테이블 셀2");
			
			xssfCell = xssfRow.createCell((short) 2);
			xssfCell.setCellStyle(cellStyle_Table_Center);
			xssfCell.setCellValue("테이블 셀3");
			*/
			
			String path = "C:/test_out/";
					
			String localFile = path + "AI_HUB_51_Worker_[" + today + "]" + ".xlsx";
			
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



