package org.bitbudget.vegetable.office.poi.test.voExportTest.bean;

import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;

@BodyVo(contentType=LayoutType.block)
public class DataBlockBean2 {

	@BodyVo(colNums="A")
	private String name;
	
	@BodyVo(colNums="B-e")
	private String age;
	
	@BodyVo(colNums="f",rowSpan=true)
	private String studentNo;
	
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
	
}
