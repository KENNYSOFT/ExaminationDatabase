package kr.KENNYSOFT.ExaminationDatabase;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertExaminationDatabase2
{
	final static String root="data_convert2";

	static WorkbookWithCellStyles[] workbooks=new WorkbookWithCellStyles[7];
	static DataValidation validation;
	static int ADJACENT_MODE=Constant.ADJACENT_MODE_WORDS;
	static int ADJACENT_VALUE=10;

	public static void main(String[] args) throws IOException
	{
		validation=Util.getDataValidation();
		for(int i=0;i<workbooks.length;++i)
		{
			String name=String.format("%d",i+2010);
			workbooks[i]=new WorkbookWithCellStyles(name,validation,ADJACENT_MODE,ADJACENT_VALUE);
		}
		findFiles(new File(root).listFiles());
		for(int i=0;i<workbooks.length;++i)
		{
			String name=String.format("%d",i+2010);
			System.out.println("ExaminationDatabase3_"+name+".xlsx");
			workbooks[i].workbook.write(new FileOutputStream("ExaminationDatabase3_"+name+".xlsx"));
		}
	}

	public static WorkbookWithCellStyles getWorkbook(int year)
	{
		try
		{
			return workbooks[year-2010];
		}
		catch(Exception e)
		{
			throw new RuntimeException();
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
				int rows=sheet.getPhysicalNumberOfRows();
				for(int i=1;i<rows;i=i+2)
				{
					if(i%10000==9999)System.out.print(".");
					XSSFRow row=sheet.getRow(i);
					XSSFRow row2=sheet.getRow(i+1);
					int year=(int)row.getCell(0).getNumericCellValue();
					String fileName=row.getCell(1).getStringCellValue();
					String prev=row.getCell(2).getStringCellValue();
					String org=row.getCell(3).getStringCellValue();
					String res=row2.getCell(3).getStringCellValue();
					String next=row.getCell(4).getStringCellValue();
					String tag=row.getCell(5).getStringCellValue();
					Util.writeRow3(getWorkbook(year),year,fileName,prev,org,res,next,tag);
				}
				workbook.close();
				System.gc();
				System.out.println(" "+(System.nanoTime()-s)/1000000+"ms");
			}
		}
	}
}