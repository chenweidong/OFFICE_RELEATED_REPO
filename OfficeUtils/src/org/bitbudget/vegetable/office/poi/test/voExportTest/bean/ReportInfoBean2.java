package org.bitbudget.vegetable.office.poi.test.voExportTest.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;

@BodyVo(contentType=LayoutType.block)
public class ReportInfoBean2 {

	@ReportInfoVo(colNums="a",alignment=AlignmentType.rigth,fontSize="10")
	private String name;
	
	@ReportInfoVo(colNums="b",alignment=AlignmentType.left,fontSize="10")
	private String age;
	
	@ReportInfoVo(colNums="c",alignment=AlignmentType.left,fontSize="10")
	private String studentNo;
	
	@ReportInfoVo(colNums="d",alignment=AlignmentType.left,fontSize="10")
	private String grade;
	
	@ReportInfoVo(colNums="e",alignment=AlignmentType.rigth,fontSize="10")
	private String leader;
	
	@ReportInfoVo(colNums="f",alignment=AlignmentType.left,fontSize="10")
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
