package org.bitbudget.vegetable.office.poi.test;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

@BodyVo(contentType=LayoutType.block)
public class AnnotationVoTest3 {

	@BodyVo(colNums="A",backColor=HSSFColor.GREY_40_PERCENT.index)
	@HeaderVo(title="学生姓名",rowNums="1-2",colNums="A"  )
	private String name;
	
	@BodyVo(colNums="B-f",alignment=AlignmentType.left,backColor=HSSFColor.GREY_40_PERCENT.index,isBold=true)
	@HeaderVo(title="学生年龄",rowNums="1-2",colNums="B",backColor=HSSFColor.YELLOW.index)
	private String age;
	
	@HeaderVo(title="学号",rowNums="2",colNums="c",backColor=HSSFColor.YELLOW.index)
	private String studentNo;
	
	@HeaderVo(title="年级",rowNums="2",colNums="d")
	private String grade;
	
	@HeaderVo(title="班主任",rowNums="2",colNums="e")
	private String leader;
	
	@HeaderVo(title="截止上年底前的成绩",rowNums="2",colNums="f")
	private String score;
	
	@HeaderVo(title="学生在校信息",rowNums="1",colNums="c-f",backColor=HSSFColor.YELLOW.index,fontColor=HSSFColor.RED.index)
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
