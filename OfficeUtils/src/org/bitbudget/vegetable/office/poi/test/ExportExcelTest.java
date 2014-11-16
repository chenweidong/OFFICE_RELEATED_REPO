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
		HSSFSheet sheet = wb.createSheet("��Ŀ�걨��");
		
//		 //�ϲ�
//		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
//		sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
//		sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
//		sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 5));
		
		
		/*
		 * ��ʽ 
		 */
		HSSFCellStyle cellStyle = wb.createCellStyle();
		//��ɫ����ɫ
		cellStyle.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		cellStyle.setFillBackgroundColor(HSSFColor.BLUE.index);
//		cellStyle.setFillPattern(HSSFCellStyle.BIG_SPOTS);
		//�߿�
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		//����
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		//cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		//����
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(HSSFColor.RED.index);
		cellStyle.setFont(font);
		//����
		cellStyle.setWrapText(true);
		
		/*
		 * ���ñ���
		 */
		HSSFRow row = sheet.createRow(0);
		HSSFRow row2 = sheet.createRow(1);
	//	row2.setHeight((short)35);
		//row.setRowStyle(cellStyle);
		//��һ��
	    HSSFCell cell = row.createCell(0);
	    cell.setCellValue("���");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(1);
	    cell.setCellValue("��Ŀ����");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(2);
	    cell.setCellValue("��λ����");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(3);
	    cell.setCellValue("��Ŀ���");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(6);
	    cell.setCellValue("�������");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(9);
	    cell.setCellValue("��Ŀ״̬");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(15);
	    cell.setCellValue("ʵʩ����");
	    cell.setCellStyle(cellStyle);
	    
	    cell = row.createCell(16);
	    cell.setCellValue("�������");
	    cell.setCellStyle(cellStyle);
	    
	    //�ڶ���
	    HSSFCell cell2 = row2.createCell(0);
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(1);
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(2);
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(3);
	    cell2.setCellValue("������ʩ");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(4);
	    cell2.setCellValue("�������");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(5);
	    cell2.setCellValue("��������");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(6);
	    cell2.setCellValue("���й���");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(7);
	    cell2.setCellValue("��������");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(8);
	    cell2.setCellValue("���й���");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(9);
	    cell2.setCellValue("����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(10);
	    cell2.setCellValue("����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(11);
	    cell2.setCellValue("����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(12);
	    cell2.setCellValue("ʵʩ");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(13);
	    cell2.setCellValue("����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(14);
	    cell2.setCellValue("�깤");
	    cell2.setCellStyle(cellStyle);
	    
	    /*
	     */
	    cell2 = row2.createCell(15);
	    cell2.setCellValue("ʵʩ����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(16);
	    cell2.setCellValue("��������");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(17);
	    cell2.setCellValue("��׼����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(18);
	    cell2.setCellValue("���úϼ�");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(19);
	    cell2.setCellValue("ǩԼ����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(20);
	    cell2.setCellValue("��ֹ����\r\n��ǰִ��");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(21);
	    cell2.setCellValue("��Ⱥ�ͬ����");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(22);
	    cell2.setCellValue("���ִ��");
	    cell2.setCellStyle(cellStyle);
	    
	    cell2 = row2.createCell(23);
	    cell2.setCellValue("Ԥ��ȫ��ִ��");
	    cell2.setCellStyle(cellStyle);
	    
	    //�ϲ�
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
