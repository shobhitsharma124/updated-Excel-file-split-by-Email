package com.getExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class split_SS {

	public String createExcelAndPutData( String str,Map<String, Object[]> data) {

		XSSFWorkbook workbook = new XSSFWorkbook(); 
        XSSFSheet sheet = workbook.createSheet("student Details"); 
        boolean b = false;
        Set<String> keyset = data.keySet(); 
        int rownum = 0; 
        for (String key : keyset) { 
            Row row = sheet.createRow(rownum++); 
            Object[] objArr = data.get(key); 
            int cellnum = 0;
            if(b==false) {
            	String [] Str = {"Test Name","First Name","Last Name","Contact	Test Name","Test Start","Test End","Score","Section Wise","Result",	"No Of Attempts","Weighted Test Score",	"Non-Compliances","Call this url to fetch session details"};
            	
            	for(String s:Str) {
            		Cell cell = row.createCell(cellnum++);
            		cell.setCellValue(s);	
            	}
            	
            	b=true;
            }else {
            for (Object obj : objArr) { 
                Cell cell = row.createCell(cellnum++); 
                if (obj instanceof String) 
                    cell.setCellValue((String)obj); 
                else if (obj instanceof Integer) 
                    cell.setCellValue((Integer)obj); 
            } 
            }
        } 
        try { 
            FileOutputStream out = new FileOutputStream(new File("D:\\New folder\\"+str+".xlsx")); 
            workbook.write(out); 
            out.close(); 
            System.out.println("written successfully."); 
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 

		
		return null;
		
	}

	public static void main(String args[]) throws Exception  
	{  
		split_SS ss = new split_SS();
	String str = "PRT score";
	int i = 0;
	
	int j=0,x=0,z=0;
	Double y = 0.0;
	File excelfile = new File("D:\\New folder\\New Microsoft Excel Worksheet.Xlsx");
	FileInputStream fis = new FileInputStream(excelfile);
	XSSFWorkbook workbook = new XSSFWorkbook(fis);
	XSSFSheet sheet = workbook.getSheetAt(0);
	Iterator<Row> rowIt = sheet.iterator();
	Map<String, Object[]> data = new TreeMap<String, Object[]>();
	while(rowIt.hasNext()) {
		Row row = rowIt.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		if(row.getCell(0).toString().contains("PRT")){
			i=i+1;
			j =(int)row.getCell(7).getNumericCellValue();
			x =(int)row.getCell(10).getNumericCellValue();
			y =new Double(row.getCell(11).getNumericCellValue());
			z = (int)row.getCell(12).getNumericCellValue();
			data.put(String.valueOf(i), new Object[]{row.getCell(0).toString(),row.getCell(1).toString(),row.getCell(3).toString(),row.getCell(4).toString(),row.getCell(5).toString(),row.getCell(6).toString(),j,row.getCell(8).toString(),row.getCell(9).toString(),x,row.getCell(11).toString(),z});
		}
		
	}
	ss.createExcelAndPutData(str, data);
	workbook.close();
	fis.close();
	}
}
