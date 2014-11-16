package org.bitbudget.vegetable.office.poi.test;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;

@BodyVo(contentType=LayoutType.block)
public class AnnotationVoTest2 {

	@ReportInfoVo(colNums="a",alignment=AlignmentType.rigth,fontSize="10")
	@BodyVo(colNums="A")
	@HeaderVo(title="ѧ������",rowNums="1-2",colNums="A"  )
	private String name;
	
	@ReportInfoVo(colNums="b",alignment=AlignmentType.left,fontSize="10")
	@BodyVo(colNums="B-e")
	@HeaderVo(title="ѧ������",rowNums="1-2",colNums="B",backColor=HSSFColor.YELLOW.index)
	private String age;
	
	@ReportInfoVo(colNums="c",alignment=AlignmentType.left,fontSize="10")
	@BodyVo(colNums="f",rowSpan=true)
	@HeaderVo(title="ѧ��",rowNums="2",colNums="c",backColor=HSSFColor.YELLOW.index)
	private String studentNo;
	
	@ReportInfoVo(colNums="d",alignment=AlignmentType.left,fontSize="10")
	@HeaderVo(title="�꼶",rowNums="2",colNums="d")
	private String grade;
	
	@ReportInfoVo(colNums="e",alignment=AlignmentType.rigth,fontSize="10")
	@HeaderVo(title="������",rowNums="2",colNums="e")
	private String leader;
	
	@ReportInfoVo(colNums="f",alignment=AlignmentType.left,fontSize="10")
	@HeaderVo(title="��ֹ�����ǰ�ĳɼ�",rowNums="2",colNums="f")
	private String score;
	
	@HeaderVo(title="ѧ����У��Ϣ",rowNums="1",colNums="c-f",backColor=HSSFColor.YELLOW.index,fontColor=HSSFColor.RED.index)
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
