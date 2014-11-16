package org.bitbudget.vegetable.office.poi.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.bitbudget.vegetable.office.poi.utils.ExportVoInfo;

import junit.framework.TestCase;


@SuppressWarnings("unchecked")
public class ExportVoInfoTest extends TestCase{

	
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
		
		List rootList = new ArrayList();
		
		//firstList
		List firstList = new ArrayList();
		List firstRowList = new ArrayList();
		
		//secondList
		List secondList = new ArrayList();

		//thirdList
		List thirdList = new ArrayList();
		List thirdBlockList = new ArrayList();
		
		/**
		 * ��ͷ
		 */
		//��һ��
		AnnotationVoTest vo1 = new AnnotationVoTest();
		vo1.setName("XX��ѧ����Ϣ��������걨����");
		firstRowList.add(vo1);
		firstList.add(firstRowList);
		
		//�ڶ���  ����
		firstRowList = new ArrayList(); 
		AnnotationVoTest2 vo2 = new AnnotationVoTest2();
		firstRowList.add(vo2);
		firstList.add(firstRowList);
		
		//������
		firstRowList = new ArrayList(); 
		vo2 = new AnnotationVoTest2();
		vo2.setName("��༶");
		vo2.setAge("zhangsanban");
		vo2.setLeader("�����");
		vo2.setScore("2013-3-3");
		firstRowList.add(vo2);
		firstList.add(firstRowList);
		
		//������  ����
		firstRowList = new ArrayList(); 
		vo2 = new AnnotationVoTest2();
		firstRowList.add(vo2);
		
		firstList.add(firstRowList);
		
		rootList.add(firstList);
		
		/**
		 * ����
		 */
		secondList.add(new AnnotationVoTest());
		rootList.add(secondList);
		
		
		/**
		 * ����
		 */
		
		//���ݿ�1
		AnnotationVoTest vo = null;
		for(int i=0;i<10;i++){
			vo = new AnnotationVoTest();
			vo.setAge("32" + i);
			vo.setName("dd" + i);
			vo.setClassInfo("clinfo" + i);
			vo.setLeader("le" + i);
			thirdBlockList.add(vo);
		}
		thirdList.add(thirdBlockList);
		
		//�ϼ�1
		thirdBlockList = new ArrayList();
		AnnotationVoTest3 v = new AnnotationVoTest3();
		v.setName("�ϼ�1");
		v.setAge("232");
		thirdBlockList.add(v);
		thirdList.add(thirdBlockList);
		
		//���ݿ�2
		thirdBlockList = new ArrayList();
		for(int i=0;i<10;i++){
			vo2 = new AnnotationVoTest2();
			vo2.setAge("32" + i);
			vo2.setName("dd" + i);
			vo2.setClassInfo("clinfo" + i);
			vo2.setLeader("le" + i);
			vo2.setStudentNo("1009");
			thirdBlockList.add(vo2);
		}
		thirdList.add(thirdBlockList);
		
		//�ϼ�2
		thirdBlockList = new ArrayList();
		AnnotationVoTest4 v4 = new AnnotationVoTest4();
		v4.setName("�ϼ�2");
		v4.setLeader("33");
		v4.setGrade("22");
		thirdBlockList.add(v4);
		thirdList.add(thirdBlockList);

		//�ϼ�3
		thirdBlockList = new ArrayList();
		AnnotationVoTest6 v6 = new AnnotationVoTest6();
		v6.setName("�ϼ�3");
		v6.setLeader("22");
		thirdBlockList.add(v6);
		thirdList.add(thirdBlockList);
		
		rootList.add(thirdList);
		
		//����
		ExportVoInfo util = new ExportVoInfo("test case");
		HSSFWorkbook workbook = util.exportExcel(rootList);
		
		FileOutputStream fos;
		try {
			fos = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/files/ExportVoInfo.xls");
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
	}
}
