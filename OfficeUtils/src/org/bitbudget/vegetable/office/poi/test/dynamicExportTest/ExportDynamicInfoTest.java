package org.bitbudget.vegetable.office.poi.test.dynamicExportTest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;
import org.bitbudget.vegetable.office.poi.style.bean.BodyStyle;
import org.bitbudget.vegetable.office.poi.style.bean.HeaderStyle;
import org.bitbudget.vegetable.office.poi.style.bean.ReportInfoStyle;
import org.bitbudget.vegetable.office.poi.utils.ExportDynamicInfo;

import junit.framework.TestCase;

public class ExportDynamicInfoTest extends TestCase {

	public void testExport(){
		
		List<List> rootList = new ArrayList<List>();
		
		/*
		 * ��ͷ
		 */
		List<List<ReportInfoStyle>> firstList = new ArrayList<List<ReportInfoStyle>>();
		List<ReportInfoStyle> firstRowList = null;
		
		ReportInfoStyle vo = null;
		
		//��һ��
		firstRowList = new ArrayList<ReportInfoStyle>();
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.center);
		vo.setBold(true);
		vo.setColNums("A-f");
		vo.setText("XX��ѧ����Ϣ��������걨����");
		
		firstRowList.add(vo);
		
		firstList.add(firstRowList);

		//�ڶ���
		firstRowList = new ArrayList<ReportInfoStyle>();
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.rigth);
		vo.setColNums("a");
		vo.setText("��༶");
		vo.setFontSize(Short.parseShort("10"));
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.left);
		vo.setColNums("b");
		vo.setFontSize(Short.parseShort("10"));
		vo.setText("������");
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.rigth);
		vo.setColNums("c");
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.rigth);
		vo.setColNums("d");
		vo.setText("");
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.left);
		vo.setColNums("e");
		vo.setFontSize(Short.parseShort("10"));
		vo.setText("�����");
		firstRowList.add(vo);
		
		vo = new ReportInfoStyle();
		vo.setAlignment(AlignmentType.left);
		vo.setColNums("f");
		vo.setFontSize(Short.parseShort("10"));
		vo.setText("2013-3-3");
		firstRowList.add(vo);
		
		firstList.add(firstRowList);
		
		
		//������--����
		firstRowList = new ArrayList<ReportInfoStyle>();
		
		vo = new ReportInfoStyle();
		vo.setColNums("a-f");
		vo.setText("");
		firstRowList.add(vo);
		
		firstList.add(firstRowList);
		
		rootList.add(firstList);
		
		/*
		 * ����
		 */
		List<HeaderStyle> secondList = new ArrayList<HeaderStyle>();
		
		HeaderStyle header = new HeaderStyle();
		header.setRowNums("4-5");
		header.setColNums("A");
		header.setTitle("ѧ������");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("4-5");
		header.setColNums("B");
		header.setTitle("ѧ������");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("c");
		header.setTitle("ѧ��");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("d");
		header.setTitle("�꼶");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("e");
		header.setTitle("������");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("5");
		header.setColNums("f");
		header.setTitle("��ֹ�����ǰ�ĳɼ�");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		secondList.add(header);
		
		header = new HeaderStyle();
		header.setRowNums("4");
		header.setColNums("c-f");
		header.setTitle("ѧ����У��Ϣ");
		header.setBackColor(HSSFColor.LIGHT_GREEN.index);
		header.setFontColor(HSSFColor.RED.index);
		secondList.add(header);
		
		rootList.add(secondList);
		
		/*
		 *	���� 
		 */
		List<List> thirdList = new ArrayList<List>();
		//�����ݣ�����
		List<List> thirdBlockList = null;
		//������
		List<BodyStyle> thirdBlockRowList = null;
		//��Ԫ��
		BodyStyle style = null;
		
		//���ݿ�1
		thirdBlockList = new ArrayList<List>();
		for(int i = 0 ; i < 10 ; i++){
			//�½�һ������
			thirdBlockRowList = new ArrayList<BodyStyle>();
			
			style = new BodyStyle();
			style.setRowSpan(true);
			style.setColNums("A");
			style.setText("dd" + i);
			
			thirdBlockRowList.add(style);
			
			style = new BodyStyle();
			style.setColNums("b-f");
			style.setText("32" + i);
			
			thirdBlockRowList.add(style);
			//���һ�м�¼
			thirdBlockList.add(thirdBlockRowList);
		}
		//��Ӵ˿�
		thirdList.add(thirdBlockList);
		
		//�ϼƿ�1
		thirdBlockList = new ArrayList<List>();
		thirdBlockRowList = new ArrayList<BodyStyle>();
		
		style = new BodyStyle();
		style.setContentType(LayoutType.block);
		style.setColNums("A");
		style.setRowSpan(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("�ϼ�");
		
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("b-f");
		style.setAlignment(AlignmentType.left);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("320");

		thirdBlockRowList.add(style);
		
		thirdBlockList.add(thirdBlockRowList);
		
		//��Ӻϼƿ�
		thirdList.add(thirdBlockList);
		
		//���ݿ�2
		thirdBlockList = new ArrayList<List>();
		for(int i = 0 ; i < 10 ; i++){
			//�½�һ������
			thirdBlockRowList = new ArrayList<BodyStyle>();
			
			style = new BodyStyle();
			style.setColNums("A");
			style.setText("dd" + i);
			
			thirdBlockRowList.add(style);
			
			style = new BodyStyle();
			style.setColNums("b-e");
			style.setText("32" + i);
			
			thirdBlockRowList.add(style);
			
			style = new BodyStyle();
			style.setColNums("f");
			style.setRowSpan(true);
			style.setText("32" + i);
			
			thirdBlockRowList.add(style);
			
			//���һ�м�¼
			thirdBlockList.add(thirdBlockRowList);
		}
		//��Ӵ˿�
		thirdList.add(thirdBlockList);
		
		//�ϼƿ�3
		thirdBlockList = new ArrayList<List>();
		thirdBlockRowList = new ArrayList<BodyStyle>();
		
		style = new BodyStyle();
		style.setColNums("A-c");
		style.setBold(true);
		style.setText("�ϼ�");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("d");
		style.setText("320");
		thirdBlockRowList.add(style);

		style = new BodyStyle();
		style.setColNums("e");
		style.setText("320");
		thirdBlockRowList.add(style);
		
		
		style = new BodyStyle();
		style.setColNums("f");
		style.setText("");
		thirdBlockRowList.add(style);
		
		thirdBlockList.add(thirdBlockRowList);
		//��Ӻϼƿ�
		thirdList.add(thirdBlockList);
		
		//�ϼƿ�4
		thirdBlockList = new ArrayList<List>();
		thirdBlockRowList = new ArrayList<BodyStyle>();
		
		style = new BodyStyle();
		style.setContentType(LayoutType.entire);
		style.setColNums("A-c");
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("�ϼ�");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("d");
		style.setBold(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("e");
		style.setAlignment(AlignmentType.center);
		style.setBold(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("22");
		thirdBlockRowList.add(style);
		
		style = new BodyStyle();
		style.setColNums("f");
		style.setBold(true);
		style.setBackColor(HSSFColor.GREY_40_PERCENT.index);
		style.setText("");
		thirdBlockRowList.add(style);
		
		thirdBlockList.add(thirdBlockRowList);
		
		//��Ӻϼƿ�
		thirdList.add(thirdBlockList);
		
		
		rootList.add(thirdList);
		
		
		//����excel
		ExportDynamicInfo util = new ExportDynamicInfo("test hh");
		HSSFWorkbook exportExcel = util.exportExcel(rootList);
		
		OutputStream io = null;
		try {
			 io = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/dynamicExportTest/exportUtils5.xls");
			exportExcel.write(io);
			io.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			if(io != null){
				try {
					io.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		} catch (IOException e){
			e.printStackTrace();
			if(io != null){
				try {
					io.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}
}
