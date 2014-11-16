package org.bitbudget.vegetable.office.poi.test;

import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;
import org.bitbudget.vegetable.office.poi.style.annotation.ReportInfoVo;

public class AnnotationVoTest5 {

	@ReportInfoVo(colNums="A-f",isBold=true)
	@HeaderVo(title="Ñ§ÉúÐÕÃû",rowNums="1-2",colNums="A"  )
	private String name;
	

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
	
}
