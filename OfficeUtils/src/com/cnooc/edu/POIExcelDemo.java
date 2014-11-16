package com.cnooc.edu;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Test;

import junit.framework.TestCase;

public class POIExcelDemo extends TestCase{

	@Test
	public void  testExportExcel(){
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("项目申报表");
		
		/*
		 * 样式 
		 */
		HSSFCellStyle cellStyle = wb.createCellStyle();
		//绿色背景色
		cellStyle.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		//边框
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		//居中
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		//cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		//字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(HSSFColor.RED.index);
		cellStyle.setFont(font);
		//换行
		cellStyle.setWrapText(true);
		
		/*
		 * 设置标题
		 */
		HSSFRow row = sheet.createRow(0);
	    //第一行
	    HSSFCell cell = row.createCell(0);
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(1);
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(2);
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(3);
	    cell.setCellValue("基础设施");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(4);
	    cell.setCellValue("管理决策");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(5);
	    cell.setCellValue("科研生产");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(6);
	    cell.setCellValue("集中管理");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(7);
	    cell.setCellValue("报批管理");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(8);
	    cell.setCellValue("自行管理");
	    cell.setCellStyle(cellStyle);
	    
	    //添加图片
	    String imgPath =  "src/com/cnooc/edu/green.png";
	    int picIndex = getPicIndex(wb,imgPath);
	    HSSFPatriarch patriarch=sheet.createDrawingPatriarch();
		HSSFClientAnchor anchor = new HSSFClientAnchor(0,0,200,200,(short)0,0,(short)0,0);
		anchor.setAnchorType(0);
		patriarch.createPicture(anchor, picIndex);	
	    
	    FileOutputStream fos;
		try {
			fos = new FileOutputStream("src/exportExcelTest.xls");
			wb.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
	}
	
	private int getPicIndex(HSSFWorkbook wb,String imgPath){
		int index = -1;
		try {
			byte[] picData = null;
			File pic = new File(imgPath );
			long length = pic.length(  );
			picData = new byte[ ( int ) length ];
			FileInputStream picIn = new FileInputStream( pic );
			picIn.read( picData );
			index = wb.addPicture( picData, HSSFWorkbook.PICTURE_TYPE_JPEG );
		} catch (IOException e) {
			e.printStackTrace();
		}  catch (Exception e) {
			e.printStackTrace();
		} 
		return index;
	}
	
	@Test
	public void testImportExcel(){
		try {
			FileInputStream fis = new FileInputStream(new File("src/exportExcelTest.xls"));
			HSSFWorkbook wb = new HSSFWorkbook(fis);
			HSSFSheet sheet = wb.getSheetAt(0);
			int lastRowNum = sheet.getLastRowNum();
//			int colsNum = 0;
			for(int i = 0 ; i < lastRowNum + 1 ;i++){
				HSSFRow row = sheet.getRow(i);
				int colsNum = row.getPhysicalNumberOfCells();
				for(int j = 0;j<colsNum ;j++){
//					int cellType = row.getCell(j).getCellType();
//					HSSFCell.CELL_TYPE_BLANK
					String cellval = row.getCell(j).getStringCellValue();
					System.out.println(i + ":" + j + " = " + cellval );
				}
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		
	}
}
