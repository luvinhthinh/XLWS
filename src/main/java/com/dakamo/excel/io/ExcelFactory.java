package com.dakamo.excel.io;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFactory {
	
	private static Map<String, File> cachedFileMap = new HashMap<String, File>();
	
	public static Workbook createWorkbookFromFile(String filename){
		try {
			File cachedFile = cachedFileMap.get(filename);
			if(cachedFile == null){
				cachedFile = new File(filename);
				cachedFileMap.put(filename, cachedFile);
			}			
			return WorkbookFactory.create(cachedFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.err.println("Can't open excel from " + filename);
		return null;
	}
	
	public static void closeWorkbook(String filename){
		File cachedFile = cachedFileMap.get(filename);
		if(cachedFile != null){
			cachedFileMap.remove(filename);
		}
	}
	
	public static void main(String[] args){
		String filename = "excels/SampleProduct.xls";
		try {
			Workbook wb1 = createWorkbookFromFile(filename);
			Workbook wb2 = createWorkbookFromFile(filename);
			
//			Sheet sheet1 = wb1.getSheet("Input");
//			Sheet sheet2 = wb2.getSheet("Input");
		        
//			System.out.println("Value from sheet 1 : " + ExcelUtil.getCellValue(sheet1, "C15"));
//			System.out.println("Value from sheet 2 : " + ExcelUtil.getCellValue(sheet2, "C15"));
//		
//			ExcelUtil.setCellValue(sheet1, "C15", "10");
//			
//			System.out.println("Value from sheet 1 : " + ExcelUtil.getCellValue(sheet1, "C15"));
//			System.out.println("Value from sheet 2 : " + ExcelUtil.getCellValue(sheet2, "C15"));
			
			wb1.close();
			wb2.close();
//			System.out.println("Value from sheet 1 : " + ExcelUtil.getCellValue(sheet1, "C15"));
//			System.out.println("Value from sheet 2 : " + ExcelUtil.getCellValue(sheet2, "C15"));
			
		} 
		  catch (Exception e) {
				e.printStackTrace();
			}
	}
}
