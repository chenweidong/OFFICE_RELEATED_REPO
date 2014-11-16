package org.bitbudget.vegetable.office.poi.test.voExportTest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.DataBlockBean1;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.DataBlockBean2;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.ReportInfoBean1;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.ReportInfoBean2;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.TitleBean;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.TotalBean1;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.TotalBean2;
import org.bitbudget.vegetable.office.poi.test.voExportTest.bean.TotalBean3;
import org.bitbudget.vegetable.office.poi.utils.ExportVoInfo;

import junit.framework.TestCase;


@SuppressWarnings("unchecked")
public class ExportVoInfoTest2 extends TestCase{

	
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
		ReportInfoBean1 bean1 = new ReportInfoBean1();
		bean1.setName("XX��ѧ����Ϣ��������걨����");
		firstRowList.add(bean1);
		firstList.add(firstRowList);
		
		//�ڶ���  ����
		firstRowList = new ArrayList(); 
		ReportInfoBean2 bean2 = new ReportInfoBean2();
		firstRowList.add(bean2);
		firstList.add(firstRowList);
		
		//������
		firstRowList = new ArrayList(); 
		bean2 = new ReportInfoBean2();
		bean2.setName("��༶");
		bean2.setAge("zhangsanban");
		bean2.setLeader("�����");
		
		bean2.setScore("2013-3-3");
		firstRowList.add(bean2);
		firstList.add(firstRowList);
		
		//������  ����
		firstRowList = new ArrayList(); 
		bean2 = new ReportInfoBean2();
		firstRowList.add(bean2);
		
		firstList.add(firstRowList);
		
		rootList.add(firstList);
		
		/**
		 * ����
		 */
		secondList.add(new TitleBean());
		rootList.add(secondList);
		
		
		/**
		 * ����
		 */
		
		//���ݿ�1
		DataBlockBean1 dataBlockBean1 = null;
		for(int i=0;i<10;i++){
			dataBlockBean1 = new DataBlockBean1();
			dataBlockBean1.setAge("32" + i);
			dataBlockBean1.setName("dd" + i);
			thirdBlockList.add(dataBlockBean1);
		}
		thirdList.add(thirdBlockList);
		
		//�ϼ�1
		thirdBlockList = new ArrayList();
		TotalBean1 totalBean1 = new TotalBean1();
		totalBean1.setName("�ϼ�1");
		totalBean1.setAge("232");
		thirdBlockList.add(totalBean1);
		thirdList.add(thirdBlockList);
		
		//���ݿ�2
		thirdBlockList = new ArrayList();
		DataBlockBean2 dataBlockBean2  = null;
		for(int i=0;i<10;i++){
			dataBlockBean2 = new DataBlockBean2();
			dataBlockBean2.setAge("32" + i);
			dataBlockBean2.setName("dd" + i);
			dataBlockBean2.setStudentNo("1009");
			thirdBlockList.add(dataBlockBean2);
		}
		thirdList.add(thirdBlockList);
		
		//�ϼ�2
		thirdBlockList = new ArrayList();
		TotalBean2 totalBean2 = new TotalBean2();
		totalBean2.setName("�ϼ�2");
		totalBean2.setLeader("33");
		totalBean2.setGrade("22");
		thirdBlockList.add(totalBean2);
		thirdList.add(thirdBlockList);

		//�ϼ�3
		thirdBlockList = new ArrayList();
		TotalBean3 totalBean3 = new TotalBean3();
		totalBean3.setName("�ϼ�3");
		totalBean3.setLeader("22");
		thirdBlockList.add(totalBean3);
		thirdList.add(thirdBlockList);
		
		rootList.add(thirdList);
		
		//����
		ExportVoInfo util = new ExportVoInfo("test case");
		HSSFWorkbook workbook = util.exportExcel(rootList);
		
		FileOutputStream fos;
		try {
			fos = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/voExportTest/ExportVoInfo.xls");
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e){
			e.printStackTrace();
		}
	}
}
