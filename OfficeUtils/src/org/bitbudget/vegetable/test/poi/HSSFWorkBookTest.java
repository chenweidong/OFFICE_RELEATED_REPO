package org.bitbudget.vegetable.test.poi;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import junit.framework.TestCase;

public class HSSFWorkBookTest extends TestCase{

	public void testGetRowTotal(){
//		HSSFWorkbook workBook = new HSSFWorkbook();
//		HSSFSheet createSheet = workBook.createSheet();
//		HSSFRow createRow = createSheet.createRow(1);
//		System.out.println("createSheet.getRow(0) == " + createSheet.getRow(0)); 
//		
//		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
//		System.out.println(lastRowNum);
//		
		//result:
		//createSheet.getRow(0) == null
		//1
//		
//		HSSFWorkbook workBook = new HSSFWorkbook();
//		HSSFSheet createSheet = workBook.createSheet();
//		HSSFRow createRow = createSheet.createRow(1);
//		System.out.println("createSheet.getRow(0) == null :" + createSheet.getRow(0) == null); //#
//		
//		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
//		System.out.println(lastRowNum);
		
		//result:
		//false
		//1

//		
//		HSSFWorkbook workBook = new HSSFWorkbook();
//		HSSFSheet createSheet = workBook.createSheet();
//		HSSFRow createRow = createSheet.createRow(0);	//#
//		System.out.println("createSheet.getRow(0) == null :" + createSheet.getRow(0) == null); 
//		
//		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
//		System.out.println(lastRowNum);
		
		//result:
		//false
		//0
		

//		HSSFWorkbook workBook = new HSSFWorkbook();
//		HSSFSheet createSheet = workBook.createSheet();
//		HSSFRow createRow0 = createSheet.createRow(0);	//#
//		HSSFRow createRow = createSheet.createRow(1);	//#
//		System.out.println("createSheet.getRow(0) == null :" + createSheet.getRow(0) == null); 
//		System.out.println("createSheet.getRow(1) == null :" + createSheet.getRow(1) == null); //#
//		
//		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
//		System.out.println(lastRowNum);
		
		//result:
		//false
		//false
		//1

		
//		HSSFWorkbook workBook = new HSSFWorkbook();
//		HSSFSheet createSheet = workBook.createSheet();
//		HSSFRow createRow0 = createSheet.createRow(0);	
//		HSSFRow createRow = createSheet.createRow(1);	
//		System.out.println("createSheet.getRow(0):" + createSheet.getRow(0) == null); //#
//		System.out.println("createSheet.getRow(1):" + createSheet.getRow(1) == null); //#
//		
//		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
//		System.out.println(lastRowNum);
		
		//result:
		//false
		//false
		//1
		
		HSSFWorkbook workBook = new HSSFWorkbook();
		HSSFSheet createSheet = workBook.createSheet();
		HSSFRow createRow0 = createSheet.createRow(0);	
		HSSFRow createRow = createSheet.createRow(1);	
		System.out.println("createSheet.getRow(0):" + createSheet.getRow(0)); //#
		System.out.println("createSheet.getRow(1):" + createSheet.getRow(1)); //#
		
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		System.out.println(lastRowNum);
		
		//result:
		//createSheet.getRow(0):org.apache.poi.hssf.usermodel.HSSFRow@1729854
		//createSheet.getRow(1):org.apache.poi.hssf.usermodel.HSSFRow@6eb38a
		//1

	}
}
