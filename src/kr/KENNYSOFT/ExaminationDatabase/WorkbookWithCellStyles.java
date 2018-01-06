package kr.KENNYSOFT.ExaminationDatabase;

import java.awt.Color;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class WorkbookWithCellStyles
{
	SXSSFWorkbook workbook;
	XSSFCellStyle cellStyleHeader;
	XSSFCellStyle cellStyleBorder;
	XSSFCellStyle cellStyleBorderRight;
	XSSFCellStyle cellStyleBorderCenter;
	XSSFCellStyle cellStyleBorderCenter2;
	XSSFCellStyle cellStyleBorderLeft;
	XSSFCellStyle cellStyleBorderUnlocked;

	public WorkbookWithCellStyles(String name,DataValidation validation,int ADJACENT_MODE,int ADJACENT_VALUE)
	{
		workbook=new SXSSFWorkbook();
		workbook.getXSSFWorkbook().getProperties().getCoreProperties().setCreator("KENNYSOFT");
		cellStyleHeader=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleHeader.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleHeader.setAlignment(HorizontalAlignment.CENTER);
		cellStyleHeader.setFillForegroundColor(new XSSFColor(new Color(141,180,226)));
		cellStyleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleHeader.setBorderTop(BorderStyle.THIN);
		cellStyleHeader.setBorderBottom(BorderStyle.THIN);
		cellStyleHeader.setBorderLeft(BorderStyle.THIN);
		cellStyleHeader.setBorderRight(BorderStyle.THIN);
		cellStyleBorder=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorder.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorder.setBorderTop(BorderStyle.THIN);
		cellStyleBorder.setBorderBottom(BorderStyle.THIN);
		cellStyleBorder.setBorderLeft(BorderStyle.THIN);
		cellStyleBorder.setBorderRight(BorderStyle.THIN);
		cellStyleBorderRight=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorderRight.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderRight.setAlignment(HorizontalAlignment.RIGHT);
		cellStyleBorderRight.setBorderTop(BorderStyle.THIN);
		cellStyleBorderRight.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderRight.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderRight.setBorderRight(BorderStyle.THIN);
		cellStyleBorderCenter=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorderCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderCenter.setAlignment(HorizontalAlignment.CENTER);
		cellStyleBorderCenter.setFillForegroundColor(new XSSFColor(new Color(217,217,217)));
		cellStyleBorderCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleBorderCenter.setBorderTop(BorderStyle.THIN);
		cellStyleBorderCenter.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderCenter.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderCenter.setBorderRight(BorderStyle.THIN);
		cellStyleBorderCenter2=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorderCenter2.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderCenter2.setAlignment(HorizontalAlignment.CENTER);
		cellStyleBorderCenter2.setBorderTop(BorderStyle.THIN);
		cellStyleBorderCenter2.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderCenter2.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderCenter2.setBorderRight(BorderStyle.THIN);
		cellStyleBorderLeft=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorderLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderLeft.setAlignment(HorizontalAlignment.LEFT);
		cellStyleBorderLeft.setBorderTop(BorderStyle.THIN);
		cellStyleBorderLeft.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderLeft.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderLeft.setBorderRight(BorderStyle.THIN);
		cellStyleBorderUnlocked=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorderUnlocked.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderUnlocked.setBorderTop(BorderStyle.THIN);
		cellStyleBorderUnlocked.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderUnlocked.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderUnlocked.setBorderRight(BorderStyle.THIN);
		cellStyleBorderUnlocked.setLocked(false);
		SXSSFSheet sheet2=workbook.createSheet(name);
		workbook.getXSSFWorkbook().getSheet(name).addIgnoredErrors(new CellRangeAddress(0,1048575,1,5),IgnoredErrorType.NUMBER_STORED_AS_TEXT);
		sheet2.setColumnWidth(1,245*32);
		sheet2.setColumnWidth(2,325*32);
		sheet2.setColumnWidth(3,245*32);
		sheet2.setColumnWidth(4,325*32);
		sheet2.setColumnWidth(5,148*32);
		sheet2.protectSheet("ExaminationDatabase3");
		sheet2.lockFormatColumns(false);
		sheet2.addValidationData(validation);
		Row row=sheet2.createRow(0);
		row.createCell(0).setCellValue("연도");
		row.createCell(1).setCellValue("파일명");
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_WORDS)row.createCell(2).setCellValue("이전 "+ADJACENT_VALUE+"어절");
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_CHARS)row.createCell(2).setCellValue("이전 "+ADJACENT_VALUE+"문자");
		row.createCell(3).setCellValue("감수 대상/결과");
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_WORDS)row.createCell(4).setCellValue("이후 "+ADJACENT_VALUE+"어절");
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_CHARS)row.createCell(4).setCellValue("이후 "+ADJACENT_VALUE+"문자");
		row.createCell(5).setCellValue("감수 유형 분류 결과");
		for(int i=0;i<6;++i)row.getCell(i).setCellStyle(cellStyleHeader);
	}
}