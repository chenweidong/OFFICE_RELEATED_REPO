package org.bitbudget.vegetable.office.poi.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.bitbudget.vegetable.office.poi.utils.history.ExportUtils2;
import org.bitbudget.vegetable.office.poi.utils.history.ExportUtils3;

import junit.framework.TestCase;


@SuppressWarnings("unchecked")
public class ExportUtils2Test extends TestCase{

	
	public void testExportHeaderAndData(){
		
		/**
		 * LIST
		 * ....list
		 * ........list
		 * ............bean1
		 * ........list
		 * ............bean1_2
		 * ....list
		 * ........bean2
		 * ....list
		 * ........list
		 * ............bean3
		 * ............bean4
		 * ............bean5
		 * ........list
		 * ............bean6
		 * ............bean7
		 * ............bean8
		 * ...
		 */
		
		List list1 = new ArrayList();
		
		List list21 = new ArrayList();
		List list211 = new ArrayList();
		List list212 = new ArrayList();
		List list213 = new ArrayList();
		
		
		List list22 = new ArrayList();

		List list23 = new ArrayList();
		List list31 = new ArrayList();
		List list32 = new ArrayList();
		List list33 = new ArrayList();
		List list34 = new ArrayList();
		List list35 = new ArrayList();
		
		AnnotationVoTest v211 = new AnnotationVoTest();
		v211.setName("XX年学生信息情况汇总申报审查表");
		list211.add(v211);
		list21.add(list211);
		
		//空行
		AnnotationVoTest2 v212 = new AnnotationVoTest2();
		list212.add(v212);
		
		v212 = new AnnotationVoTest2();
		v212.setName("填报班级");
		v212.setAge("zhangsanban");
		v212.setLeader("填报日期");
		v212.setScore("2013-3-3");
		list212.add(v212);
		
		//空行
		v212 = new AnnotationVoTest2();
		list212.add(v212);
		
		list21.add(list212);
		
		
		list22.add(new AnnotationVoTest());
		list1.add(list21);
		list1.add(list22);
		
		AnnotationVoTest vo = null;
		for(int i=0;i<10;i++){
			vo = new AnnotationVoTest();
			vo.setAge("32" + i);
			vo.setName("dd" + i);
			vo.setClassInfo("clinfo" + i);
			vo.setLeader("le" + i);
			list31.add(vo);
		}
		list23.add(list31);
		
		AnnotationVoTest3 v = new AnnotationVoTest3();
		v.setName("合计");
		v.setAge("232");
		list32.add(v);
		list23.add(list32);
		
		AnnotationVoTest2 vo2 = null;
		for(int i=0;i<10;i++){
			vo2 = new AnnotationVoTest2();
			vo2.setAge("32" + i);
			vo2.setName("dd" + i);
			vo2.setClassInfo("clinfo" + i);
			vo2.setLeader("le" + i);
			vo2.setStudentNo("1009");
			list33.add(vo2);
		}
		list23.add(list33);
		
		list1.add(list23);
		
		
		AnnotationVoTest4 v4 = new AnnotationVoTest4();
		v4.setName("合计");
		v4.setLeader("33");
		v4.setGrade("22");
		list34.add(v4);
		list23.add(list34);
		
		AnnotationVoTest6 v6 = new AnnotationVoTest6();
		v6.setName("合计");
		v6.setLeader("22");
		list35.add(v6);
		list23.add(list35);
		
		ExportUtils3 util = new ExportUtils3("test case");
		HSSFWorkbook workbook = util.exportExcel(list1);
		
		FileOutputStream fos;
		try {
			fos = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/files/exportUtils2.xls");
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
	}
}
