package com.dakamo.domain;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import com.dakamo.excel.io.ExcelFactory;

public class MyWorkBook {
	private Workbook  wb;
	public String filename;
	
	public MyWorkBook(String filename){
		this.wb = ExcelFactory.createWorkbookFromFile(filename);
	}
	
	private String getCellValue(Sheet sheet, String location){
		CellReference celRef = new CellReference(location);
        Cell cell = sheet.getRow(celRef.getRow()).getCell(celRef.getCol());
        
        switch(cell.getCellType()){
	        case Cell.CELL_TYPE_NUMERIC:
	        	System.out.println("Cell "+ location + " is numeric");
	            return new Double(cell.getNumericCellValue()).toString();
	        case Cell.CELL_TYPE_STRING:
	        	System.out.println("Cell "+ location + " is String");
	        	return cell.getStringCellValue();
	        case Cell.CELL_TYPE_FORMULA:
	        	System.out.println("Cell "+ location + " is formula");
	        	return new Double(cell.getNumericCellValue()).toString();
        }

        return "";
	}

	private void setCellValue(Sheet sheet, String location, String value){
		CellReference celRef = new CellReference(location);
        Cell cell = sheet.getRow(celRef.getRow()).getCell(celRef.getCol());
        
        switch(cell.getCellType()){
	        case Cell.CELL_TYPE_NUMERIC:
	            cell.setCellValue(new Double(value));
	        default:
	        	cell.setCellValue(value);
        }
	}
	
	public String readValueFromCell(String cell){
		String[] tokens = cell.split("!");
		return getCellValue(wb.getSheet(tokens[0]), tokens[1]);
	}
	
	public void setValueToCell(String cell, String value){
		String[] tokens = cell.split("!");
		setCellValue(wb.getSheet(tokens[0]), tokens[1], value);
	}
}
