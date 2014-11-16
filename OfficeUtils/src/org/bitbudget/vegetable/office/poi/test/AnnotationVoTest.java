package org.bitbudget.vegetable.office.poi.test;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;

@BodyVo(contentType=LayoutType.block)
public class AnnotationVoTest {

	@ReportInfoVo(colNums="A-f",isBold=true,alignment=AlignmentType.center)
	@BodyVo(colNums="A",rowSpan=true)
	@HeaderVo(title="学生姓名",rowNums="5-6",colNums="A" ,backColor=HSSFColor.LIGHT_GREEN.index )
	private String name;
	
	@BodyVo(colNums="B-f")
	@HeaderVo(title="学生年龄",rowNums="5-6",colNums="B",backColor=HSSFColor.LIGHT_GREEN.index)
	private String age;
	
	@HeaderVo(title="学号",rowNums="6",colNums="c",backColor=HSSFColor.LIGHT_GREEN.index)
	private String studentNo;
	
	@HeaderVo(title="年级",rowNums="6",colNums="d",backColor=HSSFColor.LIGHT_GREEN.index)
	private String grade;
	
	@HeaderVo(title="班主任",rowNums="6",colNums="e",backColor=HSSFColor.LIGHT_GREEN.index)
	private String leader;
	
	@HeaderVo(title="截止上年底前的成绩",rowNums="6",colNums="f",backColor=HSSFColor.LIGHT_GREEN.index)
	private String score;
	
	@HeaderVo(title="学生在校信息",rowNums="5",colNums="c-f",backColor=HSSFColor.LIGHT_GREEN.index,fontColor=HSSFColor.RED.index)
	private String classInfo;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAge() {
		return age;
	}

	public void setAge(String age) {
		this.age = age;
	}

	public String getStudentNo() {
		return studentNo;
	}

	public void setStudentNo(String studentNo) {
		this.studentNo = studentNo;
	}

	public String getGrade() {
		return grade;
	}

	public void setGrade(String grade) {
		this.grade = grade;
	}

	public String getLeader() {
		return leader;
	}

	public void setLeader(String leader) {
		this.leader = leader;
	}

	public String getScore() {
		return score;
	}

	public void setScore(String score) {
		this.score = score;
	}

	public String getClassInfo() {
		return classInfo;
	}

	public void setClassInfo(String classInfo) {
		this.classInfo = classInfo;
	}
	
}
