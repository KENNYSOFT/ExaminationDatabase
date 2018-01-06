package kr.KENNYSOFT.ExaminationDatabase;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertExaminationDatabase
{
	final static String root="data_convert";
	final static String ERROR="오류";
	final static int[][] DELETE_LIST={{2011,6312},{2011,6314},{2011,6318},{2011,6320},{2011,6352},{2011,8724},{2011,8796},{2011,12034},{2011,17592},{2011,20664},{2011,22348},{2012,16040},{2012,22002},{2012,22295},{2012,22310},{2012,23464},{2012,33170},{2012,42523},{2012,46115},{2012,69597},{2012,71279},{2012,55425},{2012,55951},{2013,88678},{2014,280},{2014,12720},{2014,12832},{2014,13184},{2014,13194},{2014,14550},{2014,21770},{2014,21858},{2014,21878},{2016,494},{2016,16272},{2016,28562},{2016,28646},{2016,29336},{2016,29340},{2016,30076},{2016,29338},{2016,38242}};

	static WorkbookWithCellStyles[][] workbooks={new WorkbookWithCellStyles[6],new WorkbookWithCellStyles[11],new WorkbookWithCellStyles[2],new WorkbookWithCellStyles[1]};
	static DataValidation validation;
	static Map<Integer,Set<Integer>> deleteSets=new HashMap<>();
	static int ADJACENT_MODE=Constant.ADJACENT_MODE_WORDS;
	static int ADJACENT_VALUE=10;

	public static void main(String[] args) throws IOException
	{
		for(int i=0;i<DELETE_LIST.length;++i)
		{
			Set<Integer> deleteSet=getDeleteSet(DELETE_LIST[i][0]);
			deleteSet.add(DELETE_LIST[i][1]);
		}
		validation=Util.getDataValidation();
		for(int i=0;i<workbooks.length;++i)
		{
			for(int j=0;j<workbooks[i].length;++j)
			{
				String name=i<workbooks.length-1?String.format("%c%d",i+'A',j+1):ERROR;
				workbooks[i][j]=new WorkbookWithCellStyles(name,validation,ADJACENT_MODE,ADJACENT_VALUE);
			}
		}
		findFiles(new File(root).listFiles());
		for(int i=0;i<workbooks.length;++i)
		{
			for(int j=0;j<workbooks[i].length;++j)
			{
				String name=i<workbooks.length-1?String.format("%c%d",i+'A',j+1):ERROR;
				System.out.println("ExaminationDatabase3_"+name+".xlsx");
				workbooks[i][j].workbook.write(new FileOutputStream("ExaminationDatabase3_"+name+".xlsx"));
			}
		}
	}

	public static Set<Integer> getDeleteSet(int year)
	{
		if(deleteSets.containsKey(year))return deleteSets.get(year);
		Set<Integer> deleteSet=new HashSet<>();
		deleteSets.put(year,deleteSet);
		return deleteSet;
	}

	public static WorkbookWithCellStyles getWorkbook(String name)
	{
		try
		{
			return workbooks[name.charAt(0)-'A'][Integer.parseInt(name.substring(1))-1];
		}
		catch(Exception e)
		{
			return workbooks[3][0];
		}
	}

	public static void findFiles(File[] files) throws IOException
	{
		for(File file : files)
		{
			if(file.isDirectory())findFiles(file.listFiles());
			else
			{
				System.out.print(file.getAbsolutePath()+" ");
				long s=System.nanoTime();
				XSSFWorkbook workbook=new XSSFWorkbook(new FileInputStream(file.getAbsolutePath()));
				XSSFSheet sheet=workbook.getSheetAt(0);
				int year=(int)sheet.getRow(1).getCell(0).getNumericCellValue();
				Set<Integer> deleteSet=getDeleteSet(year);
				int rows=sheet.getPhysicalNumberOfRows();
				for(int i=1;i<rows;i=i+2)
				{
					if(i%10000==9999)System.out.print(".");
					if(deleteSet.contains(i+1)||deleteSet.contains(i+2))continue;
					XSSFRow row=sheet.getRow(i);
					XSSFRow row2=sheet.getRow(i+1);
					String fileName=row.getCell(1).getStringCellValue();
					String prev=row.getCell(2).getStringCellValue();
					String org=row.getCell(3).getStringCellValue();
					String res=row2.getCell(3).getStringCellValue();
					String next=row.getCell(4).getStringCellValue();
					String tag=row.getCell(5).getStringCellValue().trim();
					if(res.contains("※"))
					{
						res=res.substring(0,res.indexOf("※")).trim();
						if(res.length()==0)continue;
					}
					if(org.equals(res))continue;
					if(res.contains("(?)")||res.contains("(??)"))continue;
					Util.writeRow3(getWorkbook(tag),year,fileName,prev,org,res,next,tag);
				}
				workbook.close();
				System.gc();
				System.out.println(" "+(System.nanoTime()-s)/1000000+"ms");
			}
		}
	}

	
}