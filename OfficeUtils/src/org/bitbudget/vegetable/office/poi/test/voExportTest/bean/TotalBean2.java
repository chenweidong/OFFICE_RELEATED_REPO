package org.bitbudget.vegetable.office.poi.test.voExportTest.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

@BodyVo(contentType=LayoutType.block)
public class TotalBean2 {

	@BodyVo(colNums="A-c")
	private String name;
	
	@BodyVo(colNums="d")
	private String grade;
	
	@BodyVo(colNums="e")
	private String leader;
	
	@BodyVo(colNums="f")
	private String score;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
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
