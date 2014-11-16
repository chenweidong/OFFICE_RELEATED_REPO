package org.bitbudget.vegetable.office.poi.test;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;
import org.bitbudget.vegetable.office.poi.style.annotation.HeaderVo;

@BodyVo(contentType=LayoutType.entire)
public class ExcelVoTest {

	@HeaderVo(title="ѧ������",rowNums="1-2",colNums="A")
	private String name;
	
	@HeaderVo(title="ѧ������",rowNums="1-2",colNums="B")
	private String age;
	
	@HeaderVo(title="ѧ��",rowNums="2",colNums="c")
	private String studentNo;
	
	@HeaderVo(title="�꼶",rowNums="2",colNums="d")
	private String grade;
	
	@HeaderVo(title="������",rowNums="2",colNums="e")
	private String leader;
	
	@HeaderVo(title="��ֹ�����ǰ�ĳɼ�",rowNums="2",colNums="f")
	private String score;
	
	@HeaderVo(title="ѧ����У��Ϣ",rowNums="1",colNums="c-f",backColor=HSSFColor.YELLOW.index,fontColor=HSSFColor.RED.index)
	private String classInfo;
	
}
