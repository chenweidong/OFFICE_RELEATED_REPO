package org.bitbudget.vegetable.office.poi.test.dynamicExportTest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.bean.BodyStyle;
import org.bitbudget.vegetable.office.poi.style.bean.HeaderStyle;
import org.bitbudget.vegetable.office.poi.style.bean.ReportInfoStyle;
import org.bitbudget.vegetable.office.poi.utils.ExportDynamicInfo;

import junit.framework.TestCase;

public class ExportDynamicInfoTest extends TestCase {

	public void testExport(){
		
		List<List> rootList = new ArrayList<List>();
		
		/*
		 * 表头
		 */
		List<List<ReportInfoStyle>> firstList = new ArrayList<List<ReportInfoStyle>>();
		List<ReportInfoStyle> firstRowList = null;
		
		ReportInfoStyle vo = null;
		
		//第一行
		firstRowList = new ArrayList<ReportInfoStyle>();
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.center);
		vo.setBold(true);
		vo.setColNums("A-f");
		vo.setText("XX年学生信息情况汇总申报审查表");
		
		firstRowList.add(vo);
		
		firstList.add(firstRowList);

		//第二行
		firstRowList = new ArrayList<ReportInfoStyle>();
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.rigth);
		vo.setColNums("a");
		vo.setText("填报班级");
		vo.setFontSize(Short.parseShort("10"));
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.left);
		vo.setColNums("b");
		vo.setFontSize(Short.parseShort("10"));
		vo.setText("张三版");
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.rigth);
		vo.setColNums("c");
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.rigth);
		vo.setColNums("d");
		vo.setText("");
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.left);
		vo.setColNums("e");
		vo.setFontSize(Short.parseShort("10"));
		vo.setText("填报日期");
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.left);
		vo.setColNums("f");
		vo.setFontSize(Short.parseShort("10"));
		vo.setText("2013-3-3");
		firstRowList.add(vo);
		
		firstList.add(firstRowList);
		
		
		//第三行--空行
		firstRowList = new ArrayList<ReportInfoStyle>();
		
		vo = new ReportInfoStyle();
		vo.setColNums("a-f");
		vo.setText("");
		firstRowList.add(vo);
		
		firstList.add(firstRowList);
		
		rootList.add(firstList);
		
		/*
		 * 标题
		 */
		List<HeaderStyle> secondList = new ArrayList<HeaderStyle>();
		
		HeaderStyle header = new HeaderStyle();
		header.setRowNums("4-5");
		header.setColNums("A");
		header.setTitle("学生姓名");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("4-5");
		header.setColNums("B");
		header.setTitle("学生年龄");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("c");
		header.setTitle("学号");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("d");
		header.setTitle("年级");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("e");
		header.setTitle("班主任");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("f");
		header.setTitle("截止上年底前的成绩");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("4");
		header.setColNums("c-f");
		header.setTitle("学生在校信息");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		header.setFontColor(HSSFColor.RED.index);
		secondList.add(header);
		
		rootList.add(secondList);
		
		/*
		 *	内容 
		 */
		List<List> thirdList = new ArrayList<List>();
		//块数据，跨行
		List<List> thirdBlockList = null;
		//行数据
		List<BodyStyle> thirdBlockRowList = null;
		//单元格
		BodyStyle style = null;
		
		//数据块1
		thirdBlockList = new ArrayList<List>();
		for(int i = 0 ; i < 10 ; i++){
			//新建一行数据
			thirdBlockRowList = new ArrayList<BodyStyle>();
			
			style = new BodyStyle();
			style.setRowSpan(true);
			style.setColNums("A");
			style.setText("dd" + i);
			
			thirdBlockRowList.add(style);
			
			style = new BodyStyle();
			style.setColNums("b-f");
			style.setText("32" + i);
			
			thirdBlockRowList.add(style);
			//添加一行记录
			thirdBlockList.add(thirdBlockRowList);
		}
		//添加此块
		thirdList.add(thirdBlockList);
		
		//合计块1
		thirdBlockList = new ArrayList<List>();
		thirdBlockRowList = new ArrayList<BodyStyle>();
		
		style = new BodyStyle();
		style.setContentType(LayoutType.block);
		style.setColNums("A");
		style.setRowSpan(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("合计");
		
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("b-f");
		style.setAlignment(AlignmentType.left);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("320");

		thirdBlockRowList.add(style);
		
		thirdBlockList.add(thirdBlockRowList);
		
		//添加合计块
		thirdList.add(thirdBlockList);
		
		//数据块2
		thirdBlockList = new ArrayList<List>();
		for(int i = 0 ; i < 10 ; i++){
			//新建一行数据
			thirdBlockRowList = new ArrayList<BodyStyle>();
			
			style = new BodyStyle();
			style.setColNums("A");
			style.setText("dd" + i);
			
			thirdBlockRowList.add(style);
			
			style = new BodyStyle();
			style.setColNums("b-e");
			style.setText("32" + i);
			
			thirdBlockRowList.add(style);
			
			style = new BodyStyle();
			style.setColNums("f");
			style.setRowSpan(true);
			style.setText("32" + i);
			
			thirdBlockRowList.add(style);
			
			//添加一行记录
			thirdBlockList.add(thirdBlockRowList);
		}
		//添加此块
		thirdList.add(thirdBlockList);
		
		//合计块3
		thirdBlockList = new ArrayList<List>();
		thirdBlockRowList = new ArrayList<BodyStyle>();
		
		style = new BodyStyle();
		style.setColNums("A-c");
		style.setBold(true);
		style.setText("合计");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("d");
		style.setText("320");
		thirdBlockRowList.add(style);

		style = new BodyStyle();
		style.setColNums("e");
		style.setText("320");
		thirdBlockRowList.add(style);
		
		
		style = new BodyStyle();
		style.setColNums("f");
		style.setText("");
		thirdBlockRowList.add(style);
		
		thirdBlockList.add(thirdBlockRowList);
		//添加合计块
		thirdList.add(thirdBlockList);
		
		//合计块4
		thirdBlockList = new ArrayList<List>();
		thirdBlockRowList = new ArrayList<BodyStyle>();
		
		style = new BodyStyle();
		style.setContentType(LayoutType.entire);
		style.setColNums("A-c");
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("合计");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("d");
		style.setBold(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("e");
		style.setAlignment(AlignmentType.center);
		style.setBold(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("22");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("f");
		style.setBold(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("");
		thirdBlockRowList.add(style);
		
		thirdBlockList.add(thirdBlockRowList);
		
		//添加合计块
		thirdList.add(thirdBlockList);
		
		
		rootList.add(thirdList);
		
		
		//导出excel
		ExportDynamicInfo util = new ExportDynamicInfo("test hh");
		HSSFWorkbook exportExcel = util.exportExcel(rootList);
		
		OutputStream io = null;
		try {
			 io = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/dynamicExportTest/exportUtils5.xls");
			exportExcel.write(io);
			io.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			if(io != null){
				try {
					io.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		} catch (IOException e){
			e.printStackTrace();
			if(io != null){
				try {
					io.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}
}
