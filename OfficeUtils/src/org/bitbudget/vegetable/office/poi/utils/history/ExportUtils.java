package org.bitbudget.vegetable.office.poi.utils.history;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

public class ExportUtils<T> {
	
	Class<T> clazz; 
	String sheetName;
	
	private ExportUtils(Class<T> clazz){
		this.clazz = clazz;
	}
	public ExportUtils(String sheetName){
		this.sheetName = sheetName;
	}

	public HSSFWorkbook exportExcel(Class<T> clazz){
		
		List<Field> fieldList = new ArrayList<Field>();
		
		//获取所有使用了ExcelVo注解的Field
		Field[] fields = clazz.getDeclaredFields();
		for(Field field:fields){
			if(field.isAnnotationPresent(HeaderVo.class)){
				fieldList.add(field);
			}
		}
		
		//根据相关注解生成excel标题
		HSSFWorkbook workBook = generateExcelBookWithOneSheet(this.sheetName);
		
		for(int i = 0 ; i < fieldList.size(); i++){
			Field field = fieldList.get(i);
			HeaderVo annotation = field.getAnnotation(HeaderVo.class);
			
			//标题
			String title = annotation.title();
			//列
			String colNums = annotation.colNums();
			//行
			String rowNums = annotation.rowNums();
			//背景色
			short backColor = annotation.backColor();
			//字体颜色
			short fontColor = annotation.fontColor();
			
			String[] splitRow = rowNums.split("-");
			String[] splitCols = colNums.split("-");
			
			if(splitRow.length == 1 ){
				if(splitCols.length == 1){
					genrateSheetTitle(workBook,
							Integer.parseInt(splitRow[0]),Integer.parseInt(splitRow[0]),
							getExcelColByAliases(splitCols[0]),getExcelColByAliases(splitCols[0]),
							title,backColor,fontColor);
				}else if(splitCols.length == 2){
					genrateSheetTitle(workBook,
							Integer.parseInt(splitRow[0]),Integer.parseInt(splitRow[0]),
							getExcelColByAliases(splitCols[0]),getExcelColByAliases(splitCols[1]),
							title,backColor,fontColor);
				}
			}else if(splitRow.length == 2 ){
				if(splitCols.length == 1){
					genrateSheetTitle(workBook,
							Integer.parseInt(splitRow[0]),Integer.parseInt(splitRow[1]),
							getExcelColByAliases(splitCols[0]),getExcelColByAliases(splitCols[0]),
							title,backColor,fontColor);
				}else if(splitCols.length == 2){
					genrateSheetTitle(workBook,
							Integer.parseInt(splitRow[0]),Integer.parseInt(splitRow[1]),
							getExcelColByAliases(splitCols[0]),getExcelColByAliases(splitCols[1]),
							title,backColor,fontColor);
				}
			}
		}
		  
	    FileOutputStream fos;
		try {
			fos = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/exportExcelTest1.xls");
			workBook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
		
		
		return workBook;
	}
	
	private HSSFWorkbook generateExcelBookWithOneSheet(String sheetName){
		
		HSSFWorkbook wb = new HSSFWorkbook();
		wb.createSheet(sheetName);
		
		return wb;
	}
	
	private HSSFWorkbook genrateSheetTitle(HSSFWorkbook wb,
											int firstRow,int lastRow,
											int firstCol,int lastCol,
											String title,short backColor,short fontColor){
		/*
		 * 样式 
		 */
		HSSFCellStyle cellStyle = wb.createCellStyle();
		//背景色
//		cellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
//		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFillForegroundColor(backColor);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		//边框
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		//居中
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		//字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(fontColor);
		cellStyle.setFont(font);
		//换行
		cellStyle.setWrapText(true);
		
		HSSFSheet sheet = wb.getSheetAt(0);
		
		//合并单元格
		sheet.addMergedRegion(new CellRangeAddress(firstRow -1 , lastRow -1 , firstCol, lastCol));
		//设置单元格样式
		setCellStyle(sheet,cellStyle,firstRow -1,firstCol);
		setCellStyle(sheet,cellStyle,firstRow -1,lastCol);
		setCellStyle(sheet,cellStyle,lastRow -1,firstCol);
		setCellStyle(sheet,cellStyle,lastRow -1,lastCol);
		//设置单元格内容
		HSSFRow row = sheet.getRow(firstRow -1);
		HSSFCell cell = row.getCell(firstCol);
	    cell.setCellValue(title);
		
		return wb;
	}
	
	//设置单元格样式
	private void setCellStyle(HSSFSheet sheet,CellStyle cellStyle,int rowIndex,int colIndex){
		HSSFRow row = sheet.getRow(rowIndex);
		if(row == null ){
			row = sheet.createRow(rowIndex);  
		}
		
		HSSFCell cell = row.getCell(colIndex);
		if(cell == null){
			cell = row.createCell(colIndex);
		}
		
	    cell.setCellStyle(cellStyle);
	}
	  
	private int getExcelColByAliases(String col) {
		col = col.toUpperCase();
		int count = -1;
		char[] cs = col.toCharArray();
		for (int i = 0; i < cs.length; i++) {
			count += (cs[i] - 64) * Math.pow(26, cs.length - 1 - i);
		}
		return count;
	}
}
