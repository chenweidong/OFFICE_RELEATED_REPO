package org.bitbudget.vegetable.office.poi.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import junit.framework.TestCase;


public class ExportExcelTest extends TestCase{

	
	public void testExportExcel(){
		
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("项目申报表");
		
//		 //合并
//		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
//		sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
//		sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
//		sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 5));
		
		
		/*
		 * 样式 
		 */
		HSSFCellStyle cellStyle = wb.createCellStyle();
		//绿色背景色
		cellStyle.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		cellStyle.setFillBackgroundColor(HSSFColor.BLUE.index);
//		cellStyle.setFillPattern(HSSFCellStyle.BIG_SPOTS);
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
		HSSFRow row2 = sheet.createRow(1);
	//	row2.setHeight((short)35);
		//row.setRowStyle(cellStyle);
		//第一行
	    HSSFCell cell = row.createCell(0);
	    cell.setCellValue("序号");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(1);
	    cell.setCellValue("项目名称");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(2);
	    cell.setCellValue("单位名称");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(3);
	    cell.setCellValue("项目类别");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(6);
	    cell.setCellValue("管理类别");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(9);
	    cell.setCellValue("项目状态");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(15);
	    cell.setCellValue("实施周期");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(16);
	    cell.setCellValue("费用情况");
	    cell.setCellStyle(cellStyle);
	    
	    //第二行
	    HSSFCell cell2 = row2.createCell(0);
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(1);
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(2);
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(3);
	    cell2.setCellValue("基础设施");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(4);
	    cell2.setCellValue("管理决策");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(5);
	    cell2.setCellValue("科研生产");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(6);
	    cell2.setCellValue("集中管理");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(7);
	    cell2.setCellValue("报批管理");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(8);
	    cell2.setCellValue("自行管理");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(9);
	    cell2.setCellValue("待批");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(10);
	    cell2.setCellValue("立项");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(11);
	    cell2.setCellValue("启动");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(12);
	    cell2.setCellValue("实施");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(13);
	    cell2.setCellValue("上线");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(14);
	    cell2.setCellValue("完工");
	    cell2.setCellStyle(cellStyle);
	    
	    /*
	     */
	    cell2 = row2.createCell(15);
	    cell2.setCellValue("实施周期");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(16);
	    cell2.setCellValue("待批费用");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(17);
	    cell2.setCellValue("批准费用");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(18);
	    cell2.setCellValue("费用合计");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(19);
	    cell2.setCellValue("签约费用");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(20);
	    cell2.setCellValue("截止上年\r\n底前执行");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(21);
	    cell2.setCellValue("年度合同金额额");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(22);
	    cell2.setCellValue("年度执行");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(23);
	    cell2.setCellValue("预计全年执行");
	    cell2.setCellStyle(cellStyle);
	    
	    //合并
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 5));
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 8));
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 14));
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 16, 23));
		
		
	    
	    
	    FileOutputStream fos;
		try {
			fos = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/exportExcelTest2.xls");
			wb.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
	}
}
