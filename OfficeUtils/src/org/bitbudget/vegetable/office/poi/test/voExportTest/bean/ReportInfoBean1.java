package org.bitbudget.vegetable.office.poi.test.voExportTest.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;

@BodyVo(contentType=LayoutType.block)
public class ReportInfoBean1 {

	@ReportInfoVo(colNums="A-f",isBold=true,alignment=AlignmentType.center)
	private String name;
	

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
	
}
