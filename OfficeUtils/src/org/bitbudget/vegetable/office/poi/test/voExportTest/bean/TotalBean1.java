package org.bitbudget.vegetable.office.poi.test.voExportTest.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

@BodyVo(contentType=LayoutType.block)
public class TotalBean1 {

	@BodyVo(colNums="A",backColor=HSSFColor.GREY_40_PERCENT.index)
	private String name;
	
	@BodyVo(colNums="B-f",alignment=AlignmentType.left,backColor=HSSFColor.GREY_40_PERCENT.index,isBold=true)
	private String age;
	

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
	
}
