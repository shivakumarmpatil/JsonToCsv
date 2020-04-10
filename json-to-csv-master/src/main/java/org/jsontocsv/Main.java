package org.jsontocsv;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsontocsv.parser.JSONFlattener;

public class Main {

    public static void main(String[] args) throws Exception {
        
        List<LinkedHashMap<String, String>> flatJson = JSONFlattener.parseJson(new File("files/sample.json"), "UTF-8");
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sample Excel File");
                
		int rownum = 0;
		Set<String> seenKeys = new HashSet<String>();
        for(int i = 0; i < flatJson.size(); i++) {
        	Row row = sheet.createRow(rownum++);
        	int cellnum = 0;
        	for(String key: flatJson.get(i).keySet()) {
        		String[] words = key.split("\\.");
        		String cellValue = words[words.length - 1];
        		if(!seenKeys.contains(cellValue)) {
        		Cell cell = row.createCell(cellnum++);
        		cell.setCellValue(cellValue);
        		seenKeys.add(cellValue);
        	}
        }
        }
        
        int count = 0;
        for(int i = 0; i < flatJson.size(); i++) {
        	Collection<String> currValues = flatJson.get(i).values();
        	List<String> al = new ArrayList<String>(currValues);
        	for(int r = rownum; r <= currValues.size() / seenKeys.size(); r++) {
        		Row row = sheet.createRow(r);
        		for(int c = 0; c < seenKeys.size(); c++) {
        			
        				Cell cell = row.createCell(c);
                		cell.setCellValue(al.get(count++));
        			
        		}
        	}
        }
        
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("files/sample.xlsx"));
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("sample.xlsx written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
                
        
    }

}
