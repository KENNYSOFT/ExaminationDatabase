package kr.KENNYSOFT.ExaminationDatabase;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AnalyzeExaminationDatabase
{
	final static String root="data_analyze";
	final static String title="171217";
	final static String[] tags={"오류","A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","B7","B8","B9","B10","B11","C1","C2"};

	static Map<String,Integer> tagMap=new HashMap<>();
	static Map<String,Integer> categoryMap=new HashMap<>();
	static Set<String> noCategoryFiles=new HashSet<>();
	static int[][] statYear=new int[7][];
	static int[][] statCategory=new int[9][];

	public static void main(String[] args) throws IOException
	{
		for(int i=0;i<tags.length;++i)tagMap.put(tags[i],i);
		XSSFWorkbook workbook=new XSSFWorkbook(new FileInputStream("category.xlsx"));
		for(int year=2010;year<=2016;++year)
		{
			XSSFSheet sheet=workbook.getSheet(Integer.toString(year));
			int rows=sheet.getPhysicalNumberOfRows();
			for(int i=1;i<rows;++i)
			{
				XSSFRow row=sheet.getRow(i);
				try
				{
					categoryMap.put(row.getCell(1).getStringCellValue().replace("[","_").replace("]","_").replace(",","_").trim(),(int)row.getCell(2).getNumericCellValue());
				}
				catch(Exception e)
				{
					break;
				}
			}
		}
		workbook.close();
		categoryMap.put("20130320_목포시청_김대중노벨평화상기념관_패널_문구_국어원감수.docx",2);
		categoryMap.put("20130930_국립현대미술관_[국현]과천관+텍스트_국어원감수",2);
		categoryMap.put("20140305_국회운영위원회_법령_국어원감수",3);
		categoryMap.put("20110614_몽양기념관+전시+문안(감수+요청)(국립국어원검토)",2);
		categoryMap.put("20120326_창원IAEC_재작성본1_최종전달본",2);
		categoryMap.put("20120326_창원IAEC_재작성본3_최종전달본",2);
		categoryMap.put("20120430_농진청_2012자생국화자료_최종전달본",2);
		categoryMap.put("20120507_질병관리본부_1만성콩팥병+예방과.._최종전달본",2);
		categoryMap.put("20120820_문화재청_문화재안내판+문안_최종전달본",2);
		categoryMap.put("20121218_국립현대미술관_전시장walltext_최종전달본(국어원감수)",2);
		categoryMap.put("20120326_창원IAEC_재작성본2_최종전달본",2);
		categoryMap.put("20120703_문화재청_궁중음악_최종전달본",2);
		categoryMap.put("20120703_문화재청_왕실회화_최종전달본",2);
		categoryMap.put("20121119_국립문화재연구소_발간사(일본측)_최종전달본",2);
		categoryMap.put("20121218_국립현대미술관_사진기증전_최종전달본(국어원감수)",2);
		categoryMap.put("20121218_국립현대미술관_전시기획글_+전용일_국문_최종전달본",2);
		for(int i=0;i<statYear.length;++i)statYear[i]=new int[tags.length];
		for(int i=0;i<statCategory.length;++i)statCategory[i]=new int[tags.length];
		findFiles(new File(root).listFiles());
		BufferedWriter bw=new BufferedWriter(new FileWriter(new File("stat_"+title+".csv")));
		bw.write(title+","+Arrays.toString(tags).replace("[","").replace("]","").replace(", ",","));
		bw.newLine();
		for(int i=0;i<statYear.length;++i)
		{
			bw.write((i+2010)+","+Arrays.toString(statYear[i]).replace("[","").replace("]","").replace(", ",","));
			bw.newLine();
		}
		for(int i=0;i<statCategory.length;++i)
		{
			bw.write(i+","+Arrays.toString(statCategory[i]).replace("[","").replace("]","").replace(", ",","));
			bw.newLine();
		}
		bw.close();
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
					int year=(int)row.getCell(0).getNumericCellValue();
					String fileName=row.getCell(1).getStringCellValue();
					String tag=row.getCell(5).getStringCellValue();
					int pos=tagMap.containsKey(tag)?tagMap.get(tag):0;
					int category=categoryMap.containsKey(fileName)?categoryMap.get(fileName):0;
					if(category==0)
					{
						if(!noCategoryFiles.contains(fileName))
						{
							noCategoryFiles.add(fileName);
							System.err.println(fileName);
						}
					}
					statYear[year-2010][pos]++;
					statCategory[category][pos]++;
				}
				workbook.close();
				System.gc();
				System.out.println(" "+(System.nanoTime()-s)/1000000+"ms");
			}
		}
	}
}