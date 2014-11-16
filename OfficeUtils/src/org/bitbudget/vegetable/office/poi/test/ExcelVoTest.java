package org.bitbudget.vegetable.office.poi.test;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

@BodyVo(contentType=LayoutType.entire)
public class ExcelVoTest {

	@HeaderVo(title="学生姓名",rowNums="1-2",colNums="A")
	private String name;
	
	@HeaderVo(title="学生年龄",rowNums="1-2",colNums="B")
	private String age;
	
	@HeaderVo(title="学号",rowNums="2",colNums="c")
	private String studentNo;
	
	@HeaderVo(title="年级",rowNums="2",colNums="d")
	private String grade;
	
	@HeaderVo(title="班主任",rowNums="2",colNums="e")
	private String leader;
	
	@HeaderVo(title="截止上年底前的成绩",rowNums="2",colNums="f")
	private String score;
	
	@HeaderVo(title="学生在校信息",rowNums="1",colNums="c-f",backColor=HSSFColor.YELLOW.index,fontColor=HSSFColor.RED.index)
	private String classInfo;
	
}
