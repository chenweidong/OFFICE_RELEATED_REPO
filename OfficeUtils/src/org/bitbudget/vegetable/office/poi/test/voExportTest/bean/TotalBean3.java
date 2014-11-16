package org.bitbudget.vegetable.office.poi.test.voExportTest.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

@BodyVo(contentType=LayoutType.entire)
public class TotalBean3 {

	@BodyVo(colNums="A-c",backColor=HSSFColor.GREY_40_PERCENT.index,isBold=true)
	private String name;
	
	@BodyVo(colNums="B",alignment=AlignmentType.left,backColor=HSSFColor.GREY_40_PERCENT.index,isBold=true)
	private String age;
	
	@BodyVo(colNums="c",backColor=HSSFColor.GREY_40_PERCENT.index)
	private String studentNo;
	
	@BodyVo(colNums="d",backColor=HSSFColor.GREY_40_PERCENT.index)
	private String grade;
	
	@BodyVo(colNums="e",backColor=HSSFColor.GREY_40_PERCENT.index)
	private String leader;
	
	@BodyVo(colNums="f",backColor=HSSFColor.GREY_40_PERCENT.index)
	private String score;
	
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
	
}
