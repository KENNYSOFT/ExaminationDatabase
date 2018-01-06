package kr.KENNYSOFT.ExaminationDatabase;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

public class Util
{
	public static DataValidation getDataValidation()
	{
		try
		{
			Workbook tmp=new SXSSFWorkbook();
			DataValidationHelper dvHelper=tmp.createSheet("tmp").getDataValidationHelper();
			tmp.close();
			//DataValidation validation=dvHelper.createValidation(dvHelper.createExplicitListConstraint(new String[]{"A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","C1","C2"}),new CellRangeAddressList(1,1048575,5,5));
			DataValidation validation=dvHelper.createValidation(dvHelper.createExplicitListConstraint(new String[]{"A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","B7","B8","B9","B10","B11","C1","C2"}),new CellRangeAddressList(1,1048575,5,5));
			//validation.createErrorBox("분류 표지","표기(A)\n- A1: 한글 맞춤법 위배\n- A2: 띄어쓰기 위배\n- A3: 표준어 규정 위배\n- A4: 외래어 표기법 위배\n- A5: 로마자 표기법 위배\n- A6: 문장 부호 사용법 위배\n\n문법(B)\n- B1: 부적절한 조사\n- B2: 부적절한 어미\n- B3: 부적절한 표현\n- B4: 표현의 누락\n\n어휘(C)\n- C1: 외국어‧외래어 오남용\n- C2: 부적절한 어휘(어휘 선택 오류)");
			//validation.createErrorBox("분류 표지","표기(A)\n- A1: 한글 맞춤법 위배\n- A2: 띄어쓰기 위배\n- A3: 표준어 규정 위배\n- A4: 외래어 표기법 위배\n- A5: 로마자 표기법 위배\n- A6: 문장 부호 사용법 위배\n\n문법(B)\n- B1: 부적절한 호응\n- B2: 부적절한 어순\n- B3: 부적절한 높임\n- B4: 부적절한 시제\n- B5: 부적절한 사동\n- B6: 부적절한 피동\n- B7: 부적절한 접속\n- B8: 부적절한 조사\n- B9: 부적절한 어미\n- B10: 부적절한 생략\n- B11: 부적절한 표현\n\n어휘(C)\n- C1: 외국어‧외래어 오남용\n- C2: 부적절한 어휘(어휘 선택 오류)");
			validation.createErrorBox("분류 표지","표기(A)\n- A1: 한글 맞춤법\n- A2: 띄어쓰기\n- A3: 표준어 규정\n- A4: 외래어 표기법\n- A5: 로마자 표기법\n- A6: 문장 부호 사용법\n\n문법(B)\n- B1: 호응\n- B2: 어순\n- B3: 높임\n- B4: 시제\n- B5: 사동\n- B6: 피동\n- B7: 접속\n- B8: 조사\n- B9: 어미\n- B10: 생략\n- B11: 표현\n\n어휘(C)\n- C1: 외국어‧외래어 오남용\n- C2: 어휘(어휘 선택 오류)");
			validation.setShowErrorBox(true);
			return validation;
		}
		catch(Exception e)
		{
			return null;
		}
	}
	
	public static void writeRow3(WorkbookWithCellStyles workbook3,int year,String fileName,String prev,String org,String res,String next,String tag)
	{
		SXSSFSheet sheet3=workbook3.workbook.getSheetAt(0);
		CTWorksheet ctSheet=workbook3.workbook.getXSSFWorkbook().getSheetAt(0).getCTWorksheet();
		Row row=sheet3.createRow(sheet3.getLastRowNum()+1);
		Row row2=sheet3.createRow(sheet3.getLastRowNum()+1);
		for(int i=0;i<6;++i)
		{
			XSSFCellStyle cellStyleNow=null;
			switch(i)
			{
			case 0:
			case 1:
				cellStyleNow=workbook3.cellStyleBorder;
				break;
			case 2:
				cellStyleNow=workbook3.cellStyleBorderRight;
				break;
			case 3:
				cellStyleNow=workbook3.cellStyleBorderCenter;
				break;
			case 4:
				cellStyleNow=workbook3.cellStyleBorderLeft;
				break;
			case 5:
				cellStyleNow=workbook3.cellStyleBorderUnlocked;
				break;
			}
			row.createCell(i).setCellStyle(cellStyleNow);
			row2.createCell(i).setCellStyle(cellStyleNow);
			if(i==3)row2.getCell(i).setCellStyle(workbook3.cellStyleBorderCenter2);
			else (ctSheet.isSetMergeCells()?ctSheet.getMergeCells():ctSheet.addNewMergeCells()).addNewMergeCell().setRef(new CellRangeAddress(row.getRowNum(),row2.getRowNum(),i,i).formatAsString());
		}
		row.getCell(0).setCellValue(year);
		row.getCell(1).setCellValue(fileName);
		row.getCell(2).setCellValue(prev);
		row.getCell(3).setCellValue(org);
		row2.getCell(3).setCellValue(res);
		row.getCell(4).setCellValue(next);
		row.getCell(5).setCellValue(tag);
	}
}