package org.bitbudget.vegetable.office.poi.utils.history;

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
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;

public class ExportUtils2 {
	
	String sheetName;
	HSSFCellStyle cellStyle;
	
	public ExportUtils2(String sheetName){
		this.sheetName = sheetName;
	}
	
	/**
	 * LIST
	 * ....list
	 * ........list
	 * ............bean1
	 * ........list
	 * ............bean1_2
	 * ....list
	 * ........bean2
	 * ....list
	 * ........list
	 * ............bean3
	 * ............bean4
	 * ............bean5
	 * ........list
	 * ............bean6
	 * ............bean7
	 * ............bean8
	 * ...
	 */
	public HSSFWorkbook exportExcel(List<List> list){
		
		//�������ע������excel����
		HSSFWorkbook workBook = generateExcelBookWithOneSheet(this.sheetName);
		
		//���ñ���̧ͷ
		if(list.get(0) != null && list.get(0).size() > 0){
			 exportReportInfo(workBook, list.get(0));
		}
		//���ñ������
		if(list != null && list.get(1) != null && list.get(1).get(0) != null){
			workBook = exportHeader(workBook,list.get(1).get(0).getClass());
		}
		
		//�������
		for(int i = 2 ; i < list.size() ; i++){
			
			for(int j = 0 ; j < list.get(i).size() ;j++){
				List bodyList = ((List)(list.get(i).get(j)));
				if(bodyList.get(0) != null){
					//��ȡ�����������������а�������excel��Ԫ�����ݵ��ֶ�
					List<Field> fieldList = new ArrayList<Field>();
					Field[] fields = bodyList.get(0).getClass().getDeclaredFields();
					for(Field field : fields){
						if(field.isAnnotationPresent(BodyVo.class)){
							fieldList.add(field);
						}
					}
					
					//��ȡ���ݲ�������
					LayoutType contentType = LayoutType.block;
					if(bodyList.get(0).getClass().isAnnotationPresent(BodyVo.class)){
						BodyVo annotation = bodyList.get(0).getClass().getAnnotation(BodyVo.class);
						contentType = annotation.contentType();
					}
					
					//���ݲ����������excel����
					if(contentType == LayoutType.block){
						setBlockData(workBook,bodyList,fieldList);
					}else{
						setEntireData(workBook,bodyList,fieldList);
					}
				}
				
			}
				
		}
		
		
		return workBook;
	}
	
	private HSSFWorkbook exportReportInfo(HSSFWorkbook workBook,List list){
		
		for(int i=0;i<list.size();i++){
			List subList = (List)list.get(i);
			if(subList != null && subList.size() > 0){
				//�ӵ�һ��bean�л�ȡ�洢������Ϣ���ֶ�
				Class clazz = subList.get(0).getClass();
				Field[] fields = clazz.getDeclaredFields();
				List<Field> fieldList = new ArrayList<Field>();
				for(Field field:fields){
					if(field.isAnnotationPresent(ReportInfoVo.class)){
						fieldList.add(field);
					}
				}
				//����subList�е�ÿ��bean�����ɱ�����Ϣ
				for(Object object:subList){
					
					//��ȡ��Ҫ��������к�
					int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
					if(workBook.getSheetAt(0).rowIterator().hasNext()){
						lastRowNum += 2;
					}else{
						lastRowNum += 1;
					}
					for(Field field:fieldList){
						//��ȡÿ���ֶ���ʽ
						ReportInfoVo annotation = field.getAnnotation(ReportInfoVo.class);
						AlignmentType alignment = annotation.alignment();
						short backColor = annotation.backColor();
						String colNums = annotation.colNums();
						short fontColor = annotation.fontColor();
						String fonSize = annotation.fontSize();
						boolean bold = annotation.isBold();
						//��ȡ�ֶ�ֵ
						Object value = "";
						try {
							field.setAccessible(true);
							value = field.get(object);
						} catch (IllegalArgumentException e) {
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						}
						//���ɱ�����Ϣ
						String[] colsAliases = colNums.split("-");
						int excelColNo = 0;
						if(colsAliases.length == 1){
							excelColNo = getExcelColByAliases(colsAliases[0]);
							genrateReportInfo(workBook,lastRowNum,lastRowNum,excelColNo,excelColNo,value == null?"":value.toString(),backColor,fontColor,alignment,bold,fonSize);
						}else if(colsAliases.length == 2){
							int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
							int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
							genrateReportInfo(workBook,lastRowNum,lastRowNum,excelColByAliases1,excelColByAliases2,value == null?"":value.toString(),backColor,fontColor,alignment,bold,fonSize);
							
						}else{}
					}
				}
			}
		}
		
		return workBook;
	}
	
	private HSSFWorkbook exportHeader(HSSFWorkbook workBook,Class clazz){
		
		List<Field> fieldList = new ArrayList<Field>();
		
		//��ȡ����ʹ����ExcelVoע���Field
		Field[] fields = clazz.getDeclaredFields();
		for(Field field:fields){
			if(field.isAnnotationPresent(HeaderVo.class)){
				fieldList.add(field);
			}
		}
		
		for(int i = 0 ; i < fieldList.size(); i++){
			Field field = fieldList.get(i);
			HeaderVo annotation = field.getAnnotation(HeaderVo.class);
			
			//����
			String title = annotation.title();
			//��
			String colNums = annotation.colNums();
			//��
			String rowNums = annotation.rowNums();
			//����ɫ
			short backColor = annotation.backColor();
			//������ɫ
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
		
		
		return workBook;
	}
	
	private HSSFWorkbook generateExcelBookWithOneSheet(String sheetName){
		
		HSSFWorkbook wb = new HSSFWorkbook();
		wb.createSheet(sheetName);
		
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
	private HSSFWorkbook genrateSheetTitle(HSSFWorkbook wb,
											int firstRow,int lastRow,
											int firstCol,int lastCol,
											String title,short backColor,short fontColor){
		
		genrateCell(wb,firstRow,lastRow,firstCol,lastCol,title,backColor,fontColor,
					AlignmentType.center,true,"10",true,true);
		/*
		 * ��ʽ 
		 */
//		cellStyle = wb.createCellStyle();
//		//����ɫ
////		cellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
////		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		cellStyle.setFillForegroundColor(backColor);
//		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		//�߿�
//		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
//		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
//		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
//		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
//		//����
//		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
//		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//		//����
//		HSSFFont font = wb.createFont();
//		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
//		font.setColor(fontColor);
//		cellStyle.setFont(font);
//		//����
//		cellStyle.setWrapText(true);
//		
//		HSSFSheet sheet = wb.getSheetAt(0);
//		
//		//�ϲ���Ԫ��
//		sheet.addMergedRegion(new CellRangeAddress(firstRow -1 , lastRow -1 , firstCol, lastCol));
//		//���õ�Ԫ����ʽ
////		setCellStyle(sheet,cellStyle,firstRow -1,firstCol);
////		setCellStyle(sheet,cellStyle,firstRow -1,lastCol);
////		setCellStyle(sheet,cellStyle,lastRow -1,firstCol);
////		setCellStyle(sheet,cellStyle,lastRow -1,lastCol);
//		//���õ�Ԫ����ʽ
//		for(int i = firstCol; i<= lastCol; i++){
//			for(int j = firstRow -1 ; j<= lastRow -1 ; j++){
//				setCellStyle(sheet,cellStyle,j,i);
//			}
//		}
//		//���õ�Ԫ������
//		HSSFRow row = sheet.getRow(firstRow -1);
//		HSSFCell cell = row.getCell(firstCol);
//	    cell.setCellValue(title);
		
		return wb;
	}
	
	private HSSFWorkbook genrateReportInfo(HSSFWorkbook wb,
			int firstRow,int lastRow,
			int firstCol,int lastCol,
			String title,short backColor,short fontColor,
			AlignmentType alignment,boolean isBold,String fontSize){
		
		genrateCell(wb,firstRow,lastRow,firstCol,lastCol,title,backColor,fontColor,alignment,isBold,fontSize,false,false);
//		/*
//		 * ��ʽ 
//		 */
//		cellStyle = wb.createCellStyle();
//		//����ɫ
//		cellStyle.setFillForegroundColor(backColor);
//		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		//�߿�
////		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
////		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
////		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
////		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
//		//����
//		if(AlignmentType.center == alignment){
//			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
//		}else if(AlignmentType.left == alignment){
//			cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
//		}else{
//			cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
//		}
//		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//		//����
//		HSSFFont font = wb.createFont();
//		if(isBold){
//			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
//		}
////		font.setFontHeight(Short.parseShort("240"));
//		if(fontSize != null && !"".equals(fontSize)){
//			font.setFontHeightInPoints(Short.parseShort(fontSize));
//		}else{
//			font.setFontHeightInPoints(Short.parseShort("14"));
//		}
//		font.setColor(fontColor);
//		cellStyle.setFont(font);
//		//����
////		cellStyle.setWrapText(true);
//		
//		HSSFSheet sheet = wb.getSheetAt(0);
//		
//		//�ϲ���Ԫ��
//		sheet.addMergedRegion(new CellRangeAddress(firstRow -1 , lastRow -1 , firstCol, lastCol));
//		//���õ�Ԫ����ʽ
//		for(int i = firstCol; i<= lastCol; i++){
//			for(int j = firstRow -1 ; j<= lastRow -1 ; j++){
//				setCellStyle(sheet,cellStyle,j,i);
//			}
//		}
//		//���õ�Ԫ������
//		HSSFRow row = sheet.getRow(firstRow -1);
//		HSSFCell cell = row.getCell(firstCol);
//		if(null != title){
//			cell.setCellValue(title);
//		}
//		
		return wb;
	}
	private HSSFWorkbook genrateSheetData(HSSFWorkbook wb,
			int firstRow,int lastRow,
			int firstCol,int lastCol,
			String title,short backColor,short fontColor,
			AlignmentType alignment,boolean isBold,boolean hasSolid){
		
		genrateCell(wb,firstRow,lastRow,firstCol,lastCol,title,backColor,fontColor,alignment,isBold,"10",hasSolid,false);
//		/*
//		 * ��ʽ 
//		 */
//		cellStyle = wb.createCellStyle();
//		//����ɫ
//		cellStyle.setFillForegroundColor(backColor);
//		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//		//�߿�
//		if(hasSolid){
//			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
//			cellStyle.setBorderTop(CellStyle.BORDER_THIN);
//			cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
//			cellStyle.setBorderRight(CellStyle.BORDER_THIN);
//		}
//		//����
//		if(AlignmentType.center == alignment){
//			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
//		}else if(AlignmentType.left == alignment){
//			cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
//		}else{
//			cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
//		}
//		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//		//����
//		HSSFFont font = wb.createFont();
//		if(isBold){
//			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
//		}
//		font.setColor(fontColor);
//		cellStyle.setFont(font);
//		//����
////		cellStyle.setWrapText(true);
//		
//		HSSFSheet sheet = wb.getSheetAt(0);
//		
//		//�ϲ���Ԫ��
//		sheet.addMergedRegion(new CellRangeAddress(firstRow -1 , lastRow -1 , firstCol, lastCol));
//		//���õ�Ԫ����ʽ
//		for(int i = firstCol; i<= lastCol; i++){
//			for(int j = firstRow -1 ; j<= lastRow -1 ; j++){
//				setCellStyle(sheet,cellStyle,j,i);
//			}
//		}
//		//���õ�Ԫ������
//		HSSFRow row = sheet.getRow(firstRow -1);
//		HSSFCell cell = row.getCell(firstCol);
//		if(null != title){
//			cell.setCellValue(title);
//		}
		
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
	
	private HSSFWorkbook setBlockData(HSSFWorkbook workBook,List dataList,List<Field> fieldList){
		
		//ȷ�������п�ʼ�������
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		lastRowNum += 1;
		
		//��¼��ʼ��
		int startRow = lastRowNum;
		//��¼������
		int endRow = lastRowNum;
		
		
		//�洢���е���
		List<String[]> spanList = new ArrayList<String[]>();
		
		for(int i = 0; i < dataList.size() ; i++){
			try {
				HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
				
				for(int df = 0 ; df < fieldList.size();df++){
					//��ȡ����
					Field field = fieldList.get(df);
					field.setAccessible(true);
					Object object = field.get(dataList.get(i));
					
					/*
					 * ��ȡ�洢λ�ú���ʽ
					 */
					BodyVo annotation = fieldList.get(df).getAnnotation(BodyVo.class);
					//����ɫ
					short backColor = annotation.backColor();
					//������ɫ
					short fontColor = annotation.fontColor();
					//���뷽ʽ
					AlignmentType alignment = annotation.alignment();
					//�Ƿ�Ӵ�
					boolean bold = annotation.isBold();
					
					//�����ݴ洢�ڵ�Ԫ���к�
					String colNums = annotation.colNums();
					String[] colsAliases = colNums.split("-");
					int excelColByAliases = 0;
					if(colsAliases.length == 1){
						excelColByAliases = getExcelColByAliases(colsAliases[0]);
						genrateSheetData(workBook,lastRowNum + 1,lastRowNum + 1,excelColByAliases,excelColByAliases,object == null?"":object.toString(),backColor,fontColor,alignment,bold,true);
					}else if(colsAliases.length == 2){
						int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
						int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
						genrateSheetData(workBook,lastRowNum + 1,lastRowNum + 1,excelColByAliases1,excelColByAliases2,object == null?"":object.toString(),backColor,fontColor,alignment,bold,true);
						
					}else{}
					
					//�Ƿ����
					boolean rowSpan = annotation.rowSpan();
					if(rowSpan){
						spanList.add(colsAliases);
					}
					
				}
				
				lastRowNum++;
				
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
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
			if(spanList.size() > 1){
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
		
		return workBook;
	}
	private HSSFWorkbook setEntireData(HSSFWorkbook workBook,List dataList,List<Field> fieldList){
		
		//ȷ�������п�ʼ�������
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		lastRowNum += 1;
		
		//�洢���е���
		List<String[]> spanList = new ArrayList<String[]>();
		
		for(int i = 0; i < dataList.size() ; i++){
			try {
				HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
				
				for(int df = 0 ; df < fieldList.size();df++){
					//��ȡ����
					Field field = fieldList.get(df);
					field.setAccessible(true);
					Object object = field.get(dataList.get(i));
					
					/*
					 * ��ȡ�洢λ�ú���ʽ
					 */
					BodyVo annotation = fieldList.get(df).getAnnotation(BodyVo.class);
					//����ɫ
					short backColor = annotation.backColor();
					//������ɫ
					short fontColor = annotation.fontColor();
					//���뷽ʽ
					AlignmentType alignment = annotation.alignment();
					//�Ƿ�Ӵ�
					boolean bold = annotation.isBold();
					
					//�����ݴ洢�ڵ�Ԫ���к�
					String colNums = annotation.colNums();
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
					
					//�Ƿ����
					boolean rowSpan = annotation.rowSpan();
					if(rowSpan){
						spanList.add(colsAliases);
					}
					
				}
				
				lastRowNum++;
				
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		
		return workBook;
	}
}
