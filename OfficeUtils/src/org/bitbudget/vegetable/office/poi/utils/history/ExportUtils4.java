package org.bitbudget.vegetable.office.poi.utils.history;

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
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;
import org.bitbudget.vegetable.office.poi.style.bean.ReportInfoStyle;

public class ExportUtils4 {
	
	String sheetName;
	HSSFCellStyle cellStyle;
	
	public ExportUtils4(String sheetName){
		this.sheetName = sheetName;
	}
	
	
	
	/**
	 * List
	 * 		List<Map<ReportInfoStyle,String>>
	 * 		List<HeaderStyle>
	 * 		List<Map<BodyStyle,List>>
	 */
	public HSSFWorkbook exportExcel(List<List> list){
		
		//根据相关注解生成excel标题
		HSSFWorkbook workBook = generateExcelBookWithOneSheet(this.sheetName);
		
		//设置报表抬头
		if(list.get(0) != null && list.get(0).size() > 0){
			 exportReportInfo(workBook, list.get(0));
		}
		//设置报表标题
//		if(list != null && list.get(1) != null && list.get(1).get(0) != null){
//			workBook = exportHeader(workBook,list.get(1).get(0).getClass());
//		}
		
		//填充数据
//		for(int i = 2 ; i < list.size() ; i++){
//			
//			for(int j = 0 ; j < list.get(i).size() ;j++){
//				List bodyList = ((List)(list.get(i).get(j)));
//				if(bodyList.get(0) != null){
//					//获取传入数据所属类型中包含填入excel单元格数据的字段
//					List<Field> fieldList = new ArrayList<Field>();
//					Field[] fields = bodyList.get(0).getClass().getDeclaredFields();
//					for(Field field : fields){
//						if(field.isAnnotationPresent(BodyVo.class)){
//							fieldList.add(field);
//						}
//					}
//					
//					//获取数据布局类型
//					LayoutType contentType = LayoutType.block;
//					if(bodyList.get(0).getClass().isAnnotationPresent(BodyVo.class)){
//						BodyVo annotation = bodyList.get(0).getClass().getAnnotation(BodyVo.class);
//						contentType = annotation.contentType();
//					}
//					
//					//根据布局类型填充excel数据
//					if(contentType == LayoutType.block){
//						setBlockData(workBook,bodyList,fieldList);
//					}else{
//						setEntireData(workBook,bodyList,fieldList);
//					}
//				}
//			}
//		}
		
		return workBook;
	}
	
	private HSSFWorkbook exportReportInfo(HSSFWorkbook workBook,List list){
		
		for(int i=0;i<list.size();i++){
			
			//获取需要添加数据行号
			int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
			if(workBook.getSheetAt(0).rowIterator().hasNext()){
				lastRowNum += 2;
			}else{
				lastRowNum += 1;
			}
			
			Map map = (Map)list.get(i);
			if(map != null){
				Set keySet = map.keySet();
				Iterator iterator = keySet.iterator();
				Collection values = map.values();
				Object[] vals = values.toArray();
				//遍历报表行信息
				int index = 0;
				while(iterator.hasNext()){
					ReportInfoStyle style = (ReportInfoStyle)iterator.next();
					
					
					//获取每个字段样式
					AlignmentType alignment = style.getAlignment();
					short backColor = style.getBackColor();
					String colNums = style.getColNums();
					short fontColor = style.getFontColor();
					short fonSize = style.getFontSize();
					boolean bold = style.getIsBold();
					
					
//					String reportInfo = (String)map.get(style);
					String reportInfo = vals[index] == null ? "":vals[index].toString();
					//生成报表信息
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
//					List<String> reportInfoList = (List<String>)map.get(style);
//					for(String str:reportInfoList){
//						//生成报表信息
//						String[] colsAliases = colNums.split("-");
//						int excelColNo = 0;
//						if(colsAliases.length == 1){
//							excelColNo = getExcelColByAliases(colsAliases[0]);
//							genrateReportInfo(workBook,lastRowNum,lastRowNum,excelColNo,excelColNo,str == null?"":str.toString(),backColor,fontColor,alignment,bold,fonSize);
//						}else if(colsAliases.length == 2){
//							int excelColByAliases1 = getExcelColByAliases(colsAliases[0]);
//							int excelColByAliases2 = getExcelColByAliases(colsAliases[1]);
//							genrateReportInfo(workBook,lastRowNum,lastRowNum,excelColByAliases1,excelColByAliases2,str == null?"":str.toString(),backColor,fontColor,alignment,bold,fonSize);
//							
//						}else{}
//					}
					index++;
				}
			}
		}
		
		return workBook;
	}
	
	private HSSFWorkbook exportHeader(HSSFWorkbook workBook,Class clazz){
		
		List<Field> fieldList = new ArrayList<Field>();
		
		//获取所有使用了ExcelVo注解的Field
		Field[] fields = clazz.getDeclaredFields();
		for(Field field:fields){
			if(field.isAnnotationPresent(HeaderVo.class)){
				fieldList.add(field);
			}
		}
		
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
		 * 样式 
		 */
		cellStyle = wb.createCellStyle();
		//背景色
		cellStyle.setFillForegroundColor(backColor);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		//边框
		if(hasSolid){
			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
			cellStyle.setBorderTop(CellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
			cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		}
		//居中
		if(AlignmentType.center == alignment){
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		}else if(AlignmentType.left == alignment){
			cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
		}else{
			cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
		}
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		//字体
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
		//换行
		cellStyle.setWrapText(isWrapText);
		
		HSSFSheet sheet = wb.getSheetAt(0);
		
		//合并单元格
		sheet.addMergedRegion(new CellRangeAddress(firstRow -1 , lastRow -1 , firstCol, lastCol));
		//设置单元格样式
		for(int i = firstCol; i<= lastCol; i++){
			for(int j = firstRow -1 ; j<= lastRow -1 ; j++){
				setCellStyle(sheet,cellStyle,j,i);
			}
		}
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
	
	private HSSFWorkbook setBlockData(HSSFWorkbook workBook,List dataList,List<Field> fieldList){
		
		//确定从哪行开始添加数据
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		lastRowNum += 1;
		
		//记录起始行
		int startRow = lastRowNum;
		//记录结束行
		int endRow = lastRowNum;
		
		//存储跨行的列
		List<String[]> spanList = new ArrayList<String[]>();
		
		for(int i = 0; i < dataList.size() ; i++){
			try {
				HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
				
				for(int df = 0 ; df < fieldList.size();df++){
					//获取数据
					Field field = fieldList.get(df);
					field.setAccessible(true);
					Object object = field.get(dataList.get(i));
					
					/*
					 * 获取存储位置和样式
					 */
					BodyVo annotation = fieldList.get(df).getAnnotation(BodyVo.class);
					//背景色
					short backColor = annotation.backColor();
					//字体颜色
					short fontColor = annotation.fontColor();
					//对齐方式
					AlignmentType alignment = annotation.alignment();
					//是否加粗
					boolean bold = annotation.isBold();
					
					//将数据存储在单元格列号
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
					
					//是否夸行
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
		 * 对跨行列进行跨行操作
		 */
		HSSFSheet sheet = workBook.getSheetAt(0);
		//清空原数据
		//第一个位置不清空
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
		
		//确定从哪行开始添加数据
		int lastRowNum = workBook.getSheetAt(0).getLastRowNum();
		lastRowNum += 1;
		
		//存储跨行的列
		List<String[]> spanList = new ArrayList<String[]>();
		
		for(int i = 0; i < dataList.size() ; i++){
			try {
				HSSFRow row = workBook.getSheetAt(0).createRow(lastRowNum);
				
				for(int df = 0 ; df < fieldList.size();df++){
					//获取数据
					Field field = fieldList.get(df);
					field.setAccessible(true);
					Object object = field.get(dataList.get(i));
					
					/*
					 * 获取存储位置和样式
					 */
					BodyVo annotation = fieldList.get(df).getAnnotation(BodyVo.class);
					//背景色
					short backColor = annotation.backColor();
					//字体颜色
					short fontColor = annotation.fontColor();
					//对齐方式
					AlignmentType alignment = annotation.alignment();
					//是否加粗
					boolean bold = annotation.isBold();
					
					//将数据存储在单元格列号
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
					
					//是否夸行
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
