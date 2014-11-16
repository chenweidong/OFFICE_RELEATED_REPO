package org.bitbudget.vegetable.office.poi.utils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.bean.BodyStyle;
import org.bitbudget.vegetable.office.poi.style.bean.HeaderStyle;
import org.bitbudget.vegetable.office.poi.style.bean.ReportInfoStyle;

public class ExportDynamicInfo {
	
	String sheetName;
	HSSFCellStyle cellStyle = null;
	HSSFFont font = null;
	
	public ExportDynamicInfo(String sheetName){
		this.sheetName = sheetName;
	}
	
	
	
	/**
	 * List
	 * 		List<List<ReportInfoStyle>>	---List<ReportInfoStyle>����һ�м�¼
	 * 		List<HeaderStyle>
	 * 		List<List<List<BodyStyle>>>  ----List<List<BodyStyle>>����һ������,����List<BodyStyle>����һ������
	 */
	public HSSFWorkbook exportExcel(List<List> list){
		
		//�������ע������excel����
		HSSFWorkbook workBook = generateExcelBookWithOneSheet(this.sheetName);
		
		//���ñ���̧ͷ
		if(list.get(0) != null && list.get(0).size() > 0){
			 exportReportInfo(workBook, list.get(0));
		}
		//���ñ�������
		if(list != null && list.get(1) != null && list.get(1).get(0) != null){
			workBook = exportHeader(workBook,list.get(1));
		}
		
		//�������
		for(int i = 2 ; i < list.size() ; i++){
			
			//������
			for(int j = 0 ; j < list.get(i).size() ;j++){
				//����һ������
				List<List> bodyList = ((List<List>)(list.get(i).get(j)));
				if(bodyList.get(0) != null){
					if(bodyList.get(0).get(0) != null){
						BodyStyle bodyStyle = (BodyStyle)bodyList.get(0).get(0);
						
						//��ȡ���ݲ�������
						LayoutType contentType = LayoutType.block;
						contentType = bodyStyle.getContentType();
						
						//���ݲ����������excel����
						if(contentType == LayoutType.block){
							setBlockData(workBook,bodyList);
						}else{
							setEntireData(workBook,bodyList);
						}
					
					}
				}
			}
		}
		
		return workBook;
	}
	
	/**
	 * 
	 * @param workBook
	 * @param list
	 * @return
	 */
	private HSSFWorkbook exportReportInfo(HSSFWorkbook workBook,List<ReportInfoStyle> list){
		
		for(int i=0;i<list.size();i++){
			
			//��ȡ��Ҫ���������к�
			int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
			if(workBook.getSheetAt(0).rowIterator().hasNext()){
				lastRowNum += 2;
			}else{
				lastRowNum += 1;
			}
			
			List<ReportInfoStyle> subList = (List<ReportInfoStyle>)list.get(i);
			for(ReportInfoStyle style : subList){
				//��ȡÿ���ֶ���ʽ
				AlignmentType alignment = style.getAlignment();
				short backColor = style.getBackColor();
				String colNums = style.getColNums();
				short fontColor = style.getFontColor();
				short fonSize = style.getFontSize();
				boolean bold = style.getIsBold();
				
				
				String reportInfo = style.getText();
				//���ɱ�����Ϣ
				String[] colsAliases = colNums.split("-");
				int excelColNo = 0;
				if(colsAliases.length == 1){
					excelColNo = getExcelColByAliases(colsAliases[0]);
					genrateReportInfo(workBook,lastRowNum,lastRowNum,excelColNo,excelColNo,
									reportInfo == null?"":reportInfo.toString(),
									backColor,fontColor,alignment,bold,fonSize);
				}else if(colsAliases.length == 2){
					int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
					int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
					genrateReportInfo(workBook,lastRowNum,lastRowNum,excelColByAliases1,excelColByAliases2,
										reportInfo == null?"":reportInfo.toString(),
										backColor,fontColor,alignment,bold,fonSize);
					
				}else{}
			}
		}
		
		return workBook;
	}
	
	private HSSFWorkbook exportHeader(HSSFWorkbook workBook,List<HeaderStyle> list){
		
		for(int i = 0 ; i < list.size(); i++){
			HeaderStyle headerStyle = list.get(i);
			
			//����
			String title = headerStyle.getTitle();
			//��
			String colNums = headerStyle.getColNums();
			//��
			String rowNums = headerStyle.getRowNums();
			//����ɫ
			short backColor = headerStyle.getBackColor();
			//������ɫ
			short fontColor = headerStyle.getFontColor();
			
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
		
		
		return workBook;
	}
	
	private HSSFWorkbook generateExcelBookWithOneSheet(String sheetName){
		
		HSSFWorkbook wb = new HSSFWorkbook();
		wb.createSheet(sheetName);
		
//		cellStyle = wb.createCellStyle();
//		font = wb.createFont();
		return wb;
	}
	
	
	private HSSFWorkbook genrateSheetTitle(HSSFWorkbook wb,
											int firstRow,int lastRow,
											int firstCol,int lastCol,
											String title,short backColor,short fontColor){
		
		genrateCell(wb,firstRow,lastRow,firstCol,lastCol,title,backColor,fontColor,
					AlignmentType.center,true,"10",true,true);
		return wb;
	}
	
	private HSSFWorkbook genrateReportInfo(HSSFWorkbook wb,
			int firstRow,int lastRow,
			int firstCol,int lastCol,
			String title,short backColor,short fontColor,
			AlignmentType alignment,boolean isBold,short fontSize){
		
		String fontSizeT = String.valueOf(fontSize);
		
		genrateCell(wb,firstRow,lastRow,firstCol,lastCol,title,backColor,fontColor,alignment,isBold,fontSizeT,false,false);
		return wb;
	}
	private HSSFWorkbook genrateSheetData(HSSFWorkbook wb,
			int firstRow,int lastRow,
			int firstCol,int lastCol,
			String title,short backColor,short fontColor,
			AlignmentType alignment,boolean isBold,boolean hasSolid){
		
		genrateCell(wb,firstRow,lastRow,firstCol,lastCol,title,backColor,fontColor,alignment,isBold,"10",hasSolid,false);
		
		return wb;
	}
	
	private HSSFWorkbook genrateCell(HSSFWorkbook wb,
			int firstRow,int lastRow,
			int firstCol,int lastCol,
			String title,short backColor,short fontColor,
			AlignmentType alignment,boolean isBold,String fontSize,
			boolean hasSolid,boolean isWrapText){
		/*
		 * ��ʽ 
		 */
		cellStyle = wb.createCellStyle();
		//����ɫ
		cellStyle.setFillForegroundColor(backColor);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		//�߿�
		if(hasSolid){
			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
			cellStyle.setBorderTop(CellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
			cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		}
		//����
		if(AlignmentType.center == alignment){
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		}else if(AlignmentType.left == alignment){
			cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
		}else{
			cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
		}
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		//����
		HSSFFont font = wb.createFont();
		if(isBold){
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		}
		if(fontSize != null && !"".equals(fontSize)){
			font.setFontHeightInPoints(Short.parseShort(fontSize));
		}else{
			font.setFontHeightInPoints(Short.parseShort("10"));
		}
		font.setColor(fontColor);
		cellStyle.setFont(font);
		//����
		cellStyle.setWrapText(isWrapText);
		
		HSSFSheet sheet = wb.getSheetAt(0);
		
		//�ϲ���Ԫ��
		sheet.addMergedRegion(new CellRangeAddress(firstRow -1 , lastRow -1 , firstCol, lastCol));
		//���õ�Ԫ����ʽ
		for(int i = firstCol; i<= lastCol; i++){
			for(int j = firstRow -1 ; j<= lastRow -1 ; j++){
				setCellStyle(sheet,cellStyle,j,i);
			}
		}
		//���õ�Ԫ������
		HSSFRow row = sheet.getRow(firstRow -1);
		HSSFCell cell = row.getCell(firstCol);
		cell.setCellValue(title);
		
		return wb;
	}
	
	//���õ�Ԫ����ʽ
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
	
	private HSSFWorkbook setBlockData(HSSFWorkbook workBook,List<List> bodyList){
		
		//ȷ�������п�ʼ��������
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		lastRowNum += 1;
		
		//��¼��ʼ��
		int startRow = lastRowNum;
		//��¼������
		int endRow = lastRowNum;
		
		//�洢���е���
//		List<String[]> spanList = new ArrayList<String[]>();
		List<String[]> spanList = new ArrayList<String[]>();
		if(bodyList.size() > 0 && bodyList.get(0) != null ){
			List<BodyStyle> list = bodyList.get(0);
			for(BodyStyle style : list){
				//�Ƿ����
				boolean rowSpan = style.getRowSpan();
				if(rowSpan){
					spanList.add(style.getColNums().split("-"));
				}
			}
		}
		
		try {
			HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
			
			for(int df = 0 ; df < bodyList.size();df++){
				//��ȡһ������
				List<BodyStyle> list = bodyList.get(df);
				for(BodyStyle style : list){
					/*
					 * ��ȡ�洢λ�ú���ʽ
					 */
					//����ɫ
					short backColor = style.getBackColor();
					//������ɫ
					short fontColor = style.getFontColor();
					//���뷽ʽ
					AlignmentType alignment =  style.getAlignment();
					//�Ƿ�Ӵ�
					boolean bold = style.getIsBold();
					//��ʾ��Ϣ
					String object = style.getText();
					
					//�����ݴ洢�ڵ�Ԫ���к�
					String colNums = style.getColNums();
					String[] colsAliases = colNums.split("-");
					int excelColByAliases = 0;
					if(colsAliases.length == 1){
						excelColByAliases = getExcelColByAliases(colsAliases[0]);
						genrateSheetData(workBook,
								lastRowNum + 1,
								lastRowNum + 1,
								excelColByAliases,
								excelColByAliases,object == null?"":object.toString(),
								backColor,
								fontColor,
								alignment,
								bold,
								true);
					}else if(colsAliases.length == 2){
						int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
						int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
						genrateSheetData(workBook,
								lastRowNum + 1,
								lastRowNum + 1,
								excelColByAliases1,
								excelColByAliases2,
								object == null?"":object.toString(),
								backColor,
								fontColor,
								alignment,
								bold,
								true);
						
					}else{}
					
//					//�Ƿ����
//					boolean rowSpan = style.getRowSpan();
//					if(rowSpan){
//						spanList.add(colsAliases);
//					}
				}
				lastRowNum++;
			}
			
//		for(int i = 0; i < dataList.size() ; i++){
//			try {
//				HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
//				
//				for(int df = 0 ; df < fieldList.size();df++){
//					//��ȡ����
//					Field field = fieldList.get(df);
//					field.setAccessible(true);
//					Object object = field.get(dataList.get(i));
//					
//					/*
//					 * ��ȡ�洢λ�ú���ʽ
//					 */
//					BodyVo annotation = fieldList.get(df).getAnnotation(BodyVo.class);
//					//����ɫ
//					short backColor = annotation.backColor();
//					//������ɫ
//					short fontColor = annotation.fontColor();
//					//���뷽ʽ
//					AlignmentType alignment = annotation.alignment();
//					//�Ƿ�Ӵ�
//					boolean bold = annotation.isBold();
//					
//					//�����ݴ洢�ڵ�Ԫ���к�
//					String colNums = annotation.colNums();
//					String[] colsAliases = colNums.split("-");
//					int excelColByAliases = 0;
//					if(colsAliases.length == 1){
//						excelColByAliases = getExcelColByAliases(colsAliases[0]);
//						genrateSheetData(workBook,lastRowNum + 1,lastRowNum + 1,excelColByAliases,excelColByAliases,object == null?"":object.toString(),backColor,fontColor,alignment,bold,true);
//					}else if(colsAliases.length == 2){
//						int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
//						int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
//						genrateSheetData(workBook,lastRowNum + 1,lastRowNum + 1,excelColByAliases1,excelColByAliases2,object == null?"":object.toString(),backColor,fontColor,alignment,bold,true);
//						
//					}else{}
//					
//					//�Ƿ����
//					boolean rowSpan = annotation.rowSpan();
//					if(rowSpan){
//						spanList.add(colsAliases);
//					}
//				}
//				
//				lastRowNum++;
//				
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			}
		endRow = lastRowNum - 1;
		
		/*
		 * �Կ����н��п��в���
		 */
		HSSFSheet sheet = workBook.getSheetAt(0);
		//���ԭ����
		//��һ��λ�ò����
		int isFirst = 0;
		for(int i = startRow ; i<=endRow ;i++){
			if(spanList.size() > 0){
				for(String[] cols:spanList){
					if(cols.length == 1){
						if(isFirst != 0 && i != startRow){
							sheet.getRow(i).getCell(getExcelColByAliases(cols[0])).setCellValue("");
						}
					}else if(cols.length == 2){
						for(int j = getExcelColByAliases(cols[0]);j<=getExcelColByAliases(cols[1]);j++){
							if(isFirst != 0 && i != startRow && j != getExcelColByAliases(cols[0])){
								sheet.getRow(i).getCell(getExcelColByAliases(cols[j])).setCellValue("");
							}
						}
					}
					isFirst++;
				}
			}
		}
		
		for(String[] cols:spanList){
			if(cols.length == 1){
				sheet.addMergedRegion(new CellRangeAddress(startRow,endRow,getExcelColByAliases(cols[0]),getExcelColByAliases(cols[0])));
			}else if(cols.length == 2){
				sheet.addMergedRegion(new CellRangeAddress(startRow,endRow,getExcelColByAliases(cols[0]),getExcelColByAliases(cols[1])));
			}
		}
//		
		return workBook;
	}
	private HSSFWorkbook setEntireData(HSSFWorkbook workBook,List<List> bodyList){
		
		//ȷ�������п�ʼ��������
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		lastRowNum += 1;
		
		try {
			HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
			
			for(int df = 0 ; df < bodyList.size();df++){
				//��ȡһ������
				List<BodyStyle> list = bodyList.get(df);
				for(BodyStyle style : list){
					/*
					 * ��ȡ�洢λ�ú���ʽ
					 */
					//����ɫ
					short backColor = style.getBackColor();
					//������ɫ
					short fontColor = style.getFontColor();
					//���뷽ʽ
					AlignmentType alignment =  style.getAlignment();
					//�Ƿ�Ӵ�
					boolean bold = style.getIsBold();
					//��ʾ��Ϣ
					String object = style.getText();
					
					//�����ݴ洢�ڵ�Ԫ���к�
					String colNums = style.getColNums();
					String[] colsAliases = colNums.split("-");
					int excelColByAliases = 0;
					if(colsAliases.length == 1){
						excelColByAliases = getExcelColByAliases(colsAliases[0]);
						genrateSheetData(workBook,lastRowNum + 1,lastRowNum + 1,excelColByAliases,excelColByAliases,object == null?"":object.toString(),backColor,fontColor,alignment,bold,false);
					}else if(colsAliases.length == 2){
						int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
						int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
						genrateSheetData(workBook,lastRowNum + 1,lastRowNum + 1,excelColByAliases1,excelColByAliases2,object == null?"":object.toString(),backColor,fontColor,alignment,bold,false);
						
					}else{}
					
				}
				lastRowNum++;
			}
				
				
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		}
		
		return workBook;
	}
}