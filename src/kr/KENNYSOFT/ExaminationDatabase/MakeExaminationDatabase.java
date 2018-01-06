package kr.KENNYSOFT.ExaminationDatabase;

import java.awt.Color;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import rcc.h2tlib.parser.H2TParser;
import rcc.h2tlib.parser.HWPMeta;

public class MakeExaminationDatabase
{
	//final static String root="test2";
	//final static String root_word="blank";
	final static String root="data";
	final static String root_word="data_word";
	final static int FLAG_ADD_SPACE=0x1;
	final static int FLAG_DEL_SPACE=0x2;
	final static int FLAG_DEL_PHRASE=0x4;
	final static int FLAG_REPLACE=0x8;
	final static int FLAG_ADD_PHRASE=0x10;
	final static int OPERATION_MODE_CONVERT_HWP=0x1;
	final static int OPERATION_MODE_PROCESS_TXT=0x2;
	final static int OPERATION_MODE_USE_TXT=0x4;
	final static int FILE_MODE_VERSION1=0x1;
	final static int FILE_MODE_VERSION2=0x2;
	final static int FILE_MODE_VERSION3=0x4;
	final static Pattern PATTERN_ADD_SPACE=Pattern.compile("([∨﹀˅]+)");
	final static Pattern PATTERN_DEL_SPACE=Pattern.compile("([∧^˄]+)");
	final static Pattern PATTERN_DEL_PHRASE=Pattern.compile("\\[([ \\[]*[^\\[→⟶¶]*?)([ →⟶]*[→⟶][ →⟶]*)*[※%\\* >‘∨﹀˅∧^˄]*(삭[ ]*제|뺌|없앰)[^\\]→⟶¶]*[ \\]]*\\]");
	final static Pattern PATTERN_REPLACE=Pattern.compile("\\[([ \\[]*[^\\[→⟶¶]*?)([ →⟶]*[→⟶][ →⟶]*)+([^\\]→⟶¶]*[ \\]]*)\\]");
	final static Pattern PATTERN_ADD_PHRASE=Pattern.compile("\\[([^\\[\\]¶]*[^\\[\\]¶∨﹀˅∧^˄\\x{2E80}-\\x{2EFF}\\x{31C0}-\\x{31EF}\\x{3200}-\\x{32FF}\\x{3400}-\\x{4DBF}\\x{4E00}-\\x{9FBF}\\x{F900}-\\x{FAFF}\\x{20000}-\\x{2A6DF}\\x{2F800}-\\x{2FA1F}]+[^\\[\\]¶]*)\\]");
	final static String REPLACEMENT_ADD_SPACE=null;
	final static String REPLACEMENT_DEL_SPACE="";
	final static String REPLACEMENT_DEL_PHRASE="";
	final static String REPLACEMENT_REPLACE="$3";
	final static String REPLACEMENT_ADD_PHRASE="$1";
	final static String ORIGINAL_ADD_SPACE="";
	final static String ORIGINAL_DEL_SPACE=null;
	final static String ORIGINAL_DEL_PHRASE="$1";
	final static String ORIGINAL_REPLACE="$1";
	final static String ORIGINAL_ADD_PHRASE="";
	final static Set<String> marks=new HashSet<>(Arrays.asList("·","『","』","“","”","‘","’","~","〜",",",";",":","「","」","｢","｣","《","》","～","-","ㆍ","․","/","'","\"",".","?","<",">","･","~","–","∼","(",")","[","]","〈","〉","。","‘’","…","……","○","󰡒","󰡑","󰡓","_","∙","!","•","□","“","”","...",".."));

	static H2TParser parser=new H2TParser();
	static SXSSFWorkbook workbook=new SXSSFWorkbook();
	static SXSSFWorkbook workbook2=new SXSSFWorkbook();
	static XSSFCellStyle cellStyleHeader;
	static XSSFCellStyle cellStyleBorder;
	static XSSFCellStyle cellStyleBorderUnlocked;
	static XSSFCellStyle cellStyleHeader2;
	static XSSFCellStyle cellStyleBorder2;
	static XSSFCellStyle cellStyleBorder2Right;
	static XSSFCellStyle cellStyleBorder2Center;
	static XSSFCellStyle cellStyleBorder2Center2;
	static XSSFCellStyle cellStyleBorder2Left;
	static XSSFCellStyle cellStyleBorderUnlocked2;
	static SXSSFSheet sheet;
	static Map<String,SXSSFSheet> sheet2s=new HashMap<>();
	static Map<String,CTWorksheet> ctSheets=new HashMap<>();
	static Map<String,WorkbookWithCellStyles> workbook3s=new HashMap<>();
	static DataValidation validation;
	static int FILE_MODE=FILE_MODE_VERSION1;
	static int OPERATION_MODE=OPERATION_MODE_PROCESS_TXT;
	static int ADJACENT_MODE=Constant.ADJACENT_MODE_WORDS;
	static int ADJACENT_VALUE=10;

	public static void main(String[] args) throws IOException
	{
		validation=Util.getDataValidation();
		createCellStyles();
		workbook.getXSSFWorkbook().getProperties().getCoreProperties().setCreator("KENNYSOFT");
		sheet=workbook.createSheet("Sheet1");
		workbook.getXSSFWorkbook().getSheet("Sheet1").addIgnoredErrors(new CellRangeAddress(0,1048575,1,5),IgnoredErrorType.NUMBER_STORED_AS_TEXT);
		workbook.createSheet("Sheet2");
		workbook.createSheet("Sheet3");
		sheet.setColumnWidth(1,245*32);
		sheet.setColumnWidth(2,165*32);
		sheet.setColumnWidth(3,165*32);
		sheet.setColumnWidth(4,245*32);
		sheet.setColumnWidth(5,148*32);
		sheet.protectSheet("ExaminationDatabase");
		sheet.lockFormatColumns(false);
		sheet.addValidationData(validation);
		Row row=sheet.createRow(0);
		row.createCell(0).setCellValue("연도");
		row.createCell(1).setCellValue("파일명");
		row.createCell(2).setCellValue("감수 대상");
		row.createCell(3).setCellValue("감수 결과");
		//row.createCell(4).setCellValue("문맥 정보(문장 전체)");
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_WORDS)row.createCell(4).setCellValue("문맥 정보(앞뒤 "+ADJACENT_VALUE+"어절)");
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_CHARS)row.createCell(4).setCellValue("문맥 정보(앞뒤 "+ADJACENT_VALUE+"문자)");
		row.createCell(5).setCellValue("감수 유형 분류 결과");
		for(int i=0;i<6;++i)row.getCell(i).setCellStyle(cellStyleHeader);
		workbook2.getXSSFWorkbook().getProperties().getCoreProperties().setCreator("KENNYSOFT");
		findFiles(new File(root).listFiles());
		findFiles(new File(root_word).listFiles());
		if((OPERATION_MODE&OPERATION_MODE_USE_TXT)!=0&&(FILE_MODE&FILE_MODE_VERSION1)!=0)
		{
			System.out.println("ExaminationDatabase.xlsx");
			workbook.write(new FileOutputStream("ExaminationDatabase.xlsx"));
		}
		workbook.close();
		if((OPERATION_MODE&OPERATION_MODE_USE_TXT)!=0&&(FILE_MODE&FILE_MODE_VERSION2)!=0)
		{
			System.out.println("ExaminationDatabase2.xlsx");
			workbook2.write(new FileOutputStream("ExaminationDatabase2.xlsx"));
		}
		workbook2.close();
		for(String name : workbook3s.keySet())
		{
			SXSSFWorkbook workbook3=workbook3s.get(name).workbook;
			if((OPERATION_MODE&OPERATION_MODE_USE_TXT)!=0&&(FILE_MODE&FILE_MODE_VERSION3)!=0)
			{
				System.out.println("ExaminationDatabase3_"+name+".xlsx");
				workbook3.write(new FileOutputStream("ExaminationDatabase3_"+name+".xlsx"));
			}
			workbook3.close();
		}
	}

	public static void createCellStyles()
	{
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
		cellStyleBorderUnlocked=(XSSFCellStyle)workbook.createCellStyle();
		cellStyleBorderUnlocked.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderUnlocked.setBorderTop(BorderStyle.THIN);
		cellStyleBorderUnlocked.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderUnlocked.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderUnlocked.setBorderRight(BorderStyle.THIN);
		cellStyleBorderUnlocked.setLocked(false);
		cellStyleHeader2=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleHeader2.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleHeader2.setAlignment(HorizontalAlignment.CENTER);
		cellStyleHeader2.setFillForegroundColor(new XSSFColor(new Color(141,180,226)));
		cellStyleHeader2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleHeader2.setBorderTop(BorderStyle.THIN);
		cellStyleHeader2.setBorderBottom(BorderStyle.THIN);
		cellStyleHeader2.setBorderLeft(BorderStyle.THIN);
		cellStyleHeader2.setBorderRight(BorderStyle.THIN);
		cellStyleBorder2=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleBorder2.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorder2.setBorderTop(BorderStyle.THIN);
		cellStyleBorder2.setBorderBottom(BorderStyle.THIN);
		cellStyleBorder2.setBorderLeft(BorderStyle.THIN);
		cellStyleBorder2.setBorderRight(BorderStyle.THIN);
		cellStyleBorder2Right=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleBorder2Right.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorder2Right.setAlignment(HorizontalAlignment.RIGHT);
		cellStyleBorder2Right.setBorderTop(BorderStyle.THIN);
		cellStyleBorder2Right.setBorderBottom(BorderStyle.THIN);
		cellStyleBorder2Right.setBorderLeft(BorderStyle.THIN);
		cellStyleBorder2Right.setBorderRight(BorderStyle.THIN);
		cellStyleBorder2Center=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleBorder2Center.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorder2Center.setAlignment(HorizontalAlignment.CENTER);
		cellStyleBorder2Center.setFillForegroundColor(new XSSFColor(new Color(217,217,217)));
		cellStyleBorder2Center.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleBorder2Center.setBorderTop(BorderStyle.THIN);
		cellStyleBorder2Center.setBorderBottom(BorderStyle.THIN);
		cellStyleBorder2Center.setBorderLeft(BorderStyle.THIN);
		cellStyleBorder2Center.setBorderRight(BorderStyle.THIN);
		cellStyleBorder2Center2=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleBorder2Center2.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorder2Center2.setAlignment(HorizontalAlignment.CENTER);
		cellStyleBorder2Center2.setBorderTop(BorderStyle.THIN);
		cellStyleBorder2Center2.setBorderBottom(BorderStyle.THIN);
		cellStyleBorder2Center2.setBorderLeft(BorderStyle.THIN);
		cellStyleBorder2Center2.setBorderRight(BorderStyle.THIN);
		cellStyleBorder2Left=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleBorder2Left.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorder2Left.setAlignment(HorizontalAlignment.LEFT);
		cellStyleBorder2Left.setBorderTop(BorderStyle.THIN);
		cellStyleBorder2Left.setBorderBottom(BorderStyle.THIN);
		cellStyleBorder2Left.setBorderLeft(BorderStyle.THIN);
		cellStyleBorder2Left.setBorderRight(BorderStyle.THIN);
		cellStyleBorderUnlocked2=(XSSFCellStyle)workbook2.createCellStyle();
		cellStyleBorderUnlocked2.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyleBorderUnlocked2.setBorderTop(BorderStyle.THIN);
		cellStyleBorderUnlocked2.setBorderBottom(BorderStyle.THIN);
		cellStyleBorderUnlocked2.setBorderLeft(BorderStyle.THIN);
		cellStyleBorderUnlocked2.setBorderRight(BorderStyle.THIN);
		cellStyleBorderUnlocked2.setLocked(false);
	}

	public static SXSSFSheet getSheet(String name)
	{
		if(sheet2s.containsKey(name))return sheet2s.get(name);
		SXSSFSheet sheet2=workbook2.createSheet(name);
		workbook2.getXSSFWorkbook().getSheet(name).addIgnoredErrors(new CellRangeAddress(0,1048575,1,5),IgnoredErrorType.NUMBER_STORED_AS_TEXT);
		sheet2.setColumnWidth(1,245*32);
		sheet2.setColumnWidth(2,325*32);
		sheet2.setColumnWidth(3,245*32);
		sheet2.setColumnWidth(4,325*32);
		sheet2.setColumnWidth(5,148*32);
		sheet2.protectSheet("ExaminationDatabase2");
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
		for(int i=0;i<6;++i)row.getCell(i).setCellStyle(cellStyleHeader2);
		sheet2s.put(name,sheet2);
		return sheet2;
	}

	public static SXSSFSheet getSheet(int year)
	{
		return getSheet(Integer.toString(year));
	}

	public static CTWorksheet getCTWorksheet(String name)
	{
		if(!ctSheets.containsKey(name))ctSheets.put(name,workbook2.getXSSFWorkbook().getSheet(name).getCTWorksheet());
		return ctSheets.get(name);
	}

	public static CTWorksheet getCTWorksheet(int year)
	{
		return getCTWorksheet(Integer.toString(year));
	}

	public static WorkbookWithCellStyles getWorkbook(String name)
	{
		if(workbook3s.containsKey(name))return workbook3s.get(name);
		WorkbookWithCellStyles workbook3=new WorkbookWithCellStyles(name,validation,ADJACENT_MODE,ADJACENT_VALUE);
		workbook3s.put(name,workbook3);
		return workbook3;
	}

	public static WorkbookWithCellStyles getWorkbook(int year)
	{
		return getWorkbook(Integer.toString(year));
	}

	public static void findFiles(File[] files) throws IOException
	{
		for(File file : files)
		{
			if(file.isDirectory())findFiles(file.listFiles());
			else
			{
				String fileName=file.getName();
				if((OPERATION_MODE&OPERATION_MODE_CONVERT_HWP)!=0&&fileName.endsWith(".hwp"))
				{
					System.out.println(file.getAbsolutePath());
					HWPMeta meta=new HWPMeta();
					try
					{
						parser.GetText(file.getAbsolutePath(),meta,file.getAbsolutePath().replace(".hwp",".txt"));
					}
					catch(Exception e)
					{
						System.err.println("Error occured at "+file.getAbsolutePath());
					}
				}
				if((OPERATION_MODE&OPERATION_MODE_PROCESS_TXT)!=0&&fileName.endsWith(".txt")&&!fileName.endsWith("_original.txt")&&!fileName.endsWith("_result.txt"))
				{
					System.out.println(file.getAbsolutePath());
					String hwp=new String(Files.readAllBytes(Paths.get(file.getAbsolutePath())),"UTF-8").replace("\r","¶").replace("\n","¶");
					String hwpOriginal=reprocessSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_SPACE|FLAG_DEL_PHRASE|FLAG_REPLACE|FLAG_ADD_PHRASE);
					String hwpResult=processSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_SPACE|FLAG_DEL_PHRASE|FLAG_REPLACE|FLAG_ADD_PHRASE);
					BufferedOutputStream out1=new BufferedOutputStream(new FileOutputStream(new File(file.getAbsolutePath().replace(".txt","_original.txt"))));
					out1.write(hwpOriginal.replace("¶","\r\n").getBytes());
					out1.close();
					BufferedOutputStream out2=new BufferedOutputStream(new FileOutputStream(new File(file.getAbsolutePath().replace(".txt","_result.txt"))));
					out2.write(hwpResult.replace("¶","\r\n").getBytes());
					out2.close();
				}
				if((OPERATION_MODE&OPERATION_MODE_USE_TXT)!=0&&fileName.endsWith(".txt")&&!fileName.endsWith("_original.txt")&&!fileName.endsWith("_result.txt"))
				{
					System.out.print(file.getAbsolutePath()+" ");
					long s=System.nanoTime();
					int year=Integer.parseInt(file.getName().substring(0,4));
					fileName=fileName.replace(".txt","");
					String folder=file.getParentFile().getName();
					String hwp=new String(Files.readAllBytes(Paths.get(file.getAbsolutePath())),"UTF-8").replace("\r","¶").replace("\n","¶");
					String hwpResult=processSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_SPACE|FLAG_DEL_PHRASE|FLAG_REPLACE|FLAG_ADD_PHRASE);
					String prevWords;
					String nextWords;
					Matcher matcher=PATTERN_ADD_SPACE.matcher(processSelected(hwp,FLAG_DEL_SPACE|FLAG_DEL_PHRASE|FLAG_REPLACE|FLAG_ADD_PHRASE));
					StringBuilder sb=new StringBuilder();
					//try{
					while(matcher.find())
					{
						String str=matcher.group();
						String res=new String(new char[str.length()]).replace("\0"," ");
						matcher.appendReplacement(sb,res);
						String left=getPrevWords(hwpResult,sb.length()-str.length(),1);
						String right=getNextWords(hwpResult,sb.length(),1);
						prevWords=getPrevWords(hwpResult,sb.length()-str.length()-left.length(),ADJACENT_VALUE);
						nextWords=getNextWords(hwpResult,sb.length()+right.length(),ADJACENT_VALUE);
						writeRow(year,fileName,left+right,left+res+right,prevWords+left+res+right+nextWords,"A2");
						writeRow2(folder,year,fileName,prevWords,left+right,left+res+right,nextWords,"A2");
						if((FILE_MODE&FILE_MODE_VERSION3)!=0)Util.writeRow3(getWorkbook(year),year,fileName,prevWords,left+right,left+res+right,nextWords,"A2");
					}
					System.out.print(".");
					matcher=PATTERN_DEL_SPACE.matcher(processSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_PHRASE|FLAG_REPLACE|FLAG_ADD_PHRASE));
					sb=new StringBuilder();
					while(matcher.find())
					{
						String str=matcher.group();
						String org=new String(new char[str.length()]).replace("\0"," ");
						matcher.appendReplacement(sb,REPLACEMENT_DEL_SPACE);
						String left=getPrevWords(hwpResult,sb.length(),1);
						String right=getNextWords(hwpResult,sb.length(),1);
						prevWords=getPrevWords(hwpResult,sb.length()-str.length()-left.length(),ADJACENT_VALUE);
						nextWords=getNextWords(hwpResult,sb.length()+right.length(),ADJACENT_VALUE);
						writeRow(year,fileName,left+org+right,left+right,prevWords+left+right+nextWords,"A2");
						writeRow2(folder,year,fileName,prevWords,left+org+right,left+right,nextWords,"A2");
						if((FILE_MODE&FILE_MODE_VERSION3)!=0)Util.writeRow3(getWorkbook(year),year,fileName,prevWords,left+org+right,left+right,nextWords,"A2");
					}
					System.out.print(".");
					String hwpProcessed=processSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_SPACE);
					matcher=PATTERN_DEL_PHRASE.matcher(hwpProcessed);
//					sb=new StringBuilder();
					while(matcher.find())
					{
						String str=matcher.group(1);
						String tag="B11";
//						StringBuilder sb2=new StringBuilder();
//						matcher.appendReplacement(sb2,REPLACEMENT_DEL_PHRASE);
//						sb.append(sb2);
//						String s2=sb2.toString();
//						if(s2.contains("[")||s2.contains("]"))sb=new StringBuilder(processSelected(sb.toString(),FLAG_REPLACE|FLAG_ADD_PHRASE));
						if(str.equals("ㄱ"))continue;
						int sblength=processSelected(hwpProcessed.substring(0,matcher.end()),FLAG_DEL_PHRASE|FLAG_REPLACE|FLAG_ADD_PHRASE).length();
						if(marks.contains(str.trim()))tag="A6";
						prevWords=getPrevWords(hwpResult,sblength,ADJACENT_VALUE);
						nextWords=getNextWords(hwpResult,sblength,ADJACENT_VALUE);
						writeRow(year,fileName,str,"",prevWords+nextWords,tag);
						writeRow2(folder,year,fileName,prevWords,str,"",nextWords,tag);
						if((FILE_MODE&FILE_MODE_VERSION3)!=0)Util.writeRow3(getWorkbook(year),year,fileName,prevWords,str,"",nextWords,tag);
					}
					System.out.print(".");
					hwpProcessed=processSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_SPACE|FLAG_DEL_PHRASE);
					matcher=PATTERN_REPLACE.matcher(hwpProcessed);
//					sb=new StringBuilder();
//					int posProcessed=0;
					while(matcher.find())
					{
						String org=matcher.group(1);
						String res=matcher.group(3);
						String tag="(바꾸기)";
//						StringBuilder sb2=new StringBuilder();
//						matcher.appendReplacement(sb2,REPLACEMENT_REPLACE);
//						sb.append(sb2);
//						String s2=sb2.toString();
//						if(s2.contains("[")||s2.contains("]"))
//						{
//							sb.append(processSelected(hwpProcessed.substring(posProcessed,matcher.end()),FLAG_REPLACE|FLAG_ADD_PHRASE));
//							posProcessed=matcher.end();
//						}
						if(org.equals("ㄱ")&&(res.equals("ㄴ")||res.equals("ㄴ/ㄷ")))continue;
						int sblength=processSelected(hwpProcessed.substring(0,matcher.end()),FLAG_REPLACE|FLAG_ADD_PHRASE).length();
						if(marks.contains(org.trim())&&marks.contains(res.trim()))tag="A6";
						prevWords=getPrevWords(hwpResult,sblength-res.length(),ADJACENT_VALUE);
						nextWords=getNextWords(hwpResult,sblength,ADJACENT_VALUE);
						writeRow(year,fileName,org,res,prevWords+res+nextWords,tag);
						writeRow2(folder,year,fileName,prevWords,org,res,nextWords,tag);
						if((FILE_MODE&FILE_MODE_VERSION3)!=0)Util.writeRow3(getWorkbook(year),year,fileName,prevWords,org,res,nextWords,tag);
//						for(String resp : res.split("/"))
//						{
//							writeRow(year,fileName,org,res,prevWords+resp+nextWords,tag);
//							writeRow2(folder,year,fileName,prevWords,org,resp,nextWords,tag);
//						}
					}
					System.out.print(".");
					matcher=PATTERN_ADD_PHRASE.matcher(processSelected(hwp,FLAG_ADD_SPACE|FLAG_DEL_SPACE|FLAG_DEL_PHRASE|FLAG_REPLACE));
					sb=new StringBuilder();
					while(matcher.find())
					{
						String str=matcher.group(1);
						String tag="B10";
						matcher.appendReplacement(sb,REPLACEMENT_ADD_PHRASE);
						if(str.equals("ㄱ")||str.trim().length()==0)continue;
						if(marks.contains(str.trim()))tag="A6";
						prevWords=getPrevWords(hwpResult,sb.length()-str.length(),ADJACENT_VALUE);
						nextWords=getNextWords(hwpResult,sb.length(),ADJACENT_VALUE);
						writeRow(year,fileName,"",str,prevWords+str+nextWords,tag);
						writeRow2(folder,year,fileName,prevWords,"",str,nextWords,tag);
						if((FILE_MODE&FILE_MODE_VERSION3)!=0)Util.writeRow3(getWorkbook(year),year,fileName,prevWords,"",str,nextWords,tag);
					}
					System.out.println(" "+(System.nanoTime()-s)/1000000+"ms");
					//}catch(Exception e){System.out.println(sb.toString());System.err.println(hwpResult);e.printStackTrace();throw new RuntimeException();}
				}
			}
		}
	}

	public static String getPrevWords(String hwp,int pos,int n)
	{
		if(pos<0)pos=0;
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_WORDS)
		{
			int cut;
			int state=0;
			int cnt=0;
			for(cut=pos-1;cut>=0;cut--)
			{
				if(state==0&&!(hwp.charAt(cut)==' '||hwp.charAt(cut)=='¶'))state=1;
				if(state==1&&(hwp.charAt(cut)==' '||hwp.charAt(cut)=='¶'))
				{
					state=0;
					cnt++;
				}
				if(cnt==n)
				{
					cut++;
					break;
				}
			}
			if(cut<0)cut=0;
			return hwp.substring(cut,pos);
		}
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_CHARS)return hwp.substring(Math.max(pos-ADJACENT_VALUE,0),pos);
		return null;
	}

	public static String getNextWords(String hwp,int pos,int n)
	{
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_WORDS)
		{
			int cut;
			int state=0;
			int cnt=0;
			for(cut=pos;cut<hwp.length();cut++)
			{
				if(state==0&&!(hwp.charAt(cut)==' '||hwp.charAt(cut)=='¶'))state=1;
				if(state==1&&(hwp.charAt(cut)==' '||hwp.charAt(cut)=='¶'))
				{
					state=0;
					cnt++;
				}
				if(cnt==n)break;
			}
			return hwp.substring(pos,cut);
		}
		if(ADJACENT_MODE==Constant.ADJACENT_MODE_CHARS)return hwp.substring(pos,Math.min(pos+ADJACENT_VALUE,hwp.length()-1));
		return null;
	}

	public static void writeRow(int year,String fileName,String org,String res,String full,String tag)
	{
		if((FILE_MODE&FILE_MODE_VERSION1)!=0)
		{
			Row row=sheet.createRow(sheet.getLastRowNum()+1);
			row.createCell(0).setCellValue(year);
			row.createCell(1).setCellValue(fileName);
			row.createCell(2).setCellValue(org);
			row.createCell(3).setCellValue(res);
			row.createCell(4).setCellValue(full);
			row.createCell(5).setCellValue(tag);
			for(int i=0;i<6;++i)row.getCell(i).setCellStyle(cellStyleBorder);
			row.getCell(5).setCellStyle(cellStyleBorderUnlocked);
		}
	}

	public static void writeRow2(String folder,int year,String fileName,String prev,String org,String res,String next,String tag)
	{
		if((FILE_MODE&FILE_MODE_VERSION2)!=0)
		{
			//SXSSFSheet sheet2=getSheet(folder);
			SXSSFSheet sheet2=getSheet(year);
			//CTWorksheet ctSheet=getCTWorksheet(folder);
			CTWorksheet ctSheet=getCTWorksheet(year);
			Row row=sheet2.createRow(sheet2.getLastRowNum()+1);
			Row row2=sheet2.createRow(sheet2.getLastRowNum()+1);
			for(int i=0;i<6;++i)
			{
				XSSFCellStyle cellStyleNow=null;
				switch(i)
				{
				case 0:
				case 1:
					cellStyleNow=cellStyleBorder2;
					break;
				case 2:
					cellStyleNow=cellStyleBorder2Right;
					break;
				case 3:
					cellStyleNow=cellStyleBorder2Center;
					break;
				case 4:
					cellStyleNow=cellStyleBorder2Left;
					break;
				case 5:
					cellStyleNow=cellStyleBorderUnlocked2;
					break;
				}
				row.createCell(i).setCellStyle(cellStyleNow);
				row2.createCell(i).setCellStyle(cellStyleNow);
				if(i==3)row2.getCell(i).setCellStyle(cellStyleBorder2Center2);
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

	public static String processSelected(String hwp,int flag)
	{
		if((flag&FLAG_DEL_PHRASE)!=0)hwp=PATTERN_DEL_PHRASE.matcher(hwp).replaceAll(REPLACEMENT_DEL_PHRASE);
		if((flag&FLAG_REPLACE)!=0)hwp=PATTERN_REPLACE.matcher(hwp).replaceAll(REPLACEMENT_REPLACE);
		if((flag&FLAG_ADD_PHRASE)!=0)hwp=PATTERN_ADD_PHRASE.matcher(hwp).replaceAll(REPLACEMENT_ADD_PHRASE);
		if((flag&FLAG_ADD_SPACE)!=0)hwp=hwp.replaceAll("[∨﹀˅]"," ");
		if((flag&FLAG_DEL_SPACE)!=0)hwp=PATTERN_DEL_SPACE.matcher(hwp).replaceAll(REPLACEMENT_DEL_SPACE);
		return hwp;
	}

	public static String reprocessSelected(String hwp,int flag)
	{
		if((flag&FLAG_DEL_PHRASE)!=0)hwp=PATTERN_DEL_PHRASE.matcher(hwp).replaceAll(ORIGINAL_DEL_PHRASE);
		if((flag&FLAG_REPLACE)!=0)hwp=PATTERN_REPLACE.matcher(hwp).replaceAll(ORIGINAL_REPLACE);
		if((flag&FLAG_ADD_PHRASE)!=0)hwp=PATTERN_ADD_PHRASE.matcher(hwp).replaceAll(ORIGINAL_ADD_PHRASE);
		if((flag&FLAG_ADD_SPACE)!=0)hwp=PATTERN_ADD_SPACE.matcher(hwp).replaceAll(ORIGINAL_ADD_SPACE);
		if((flag&FLAG_DEL_SPACE)!=0)hwp=hwp.replaceAll("[∧^˄]"," ");
		return hwp;
	}
}