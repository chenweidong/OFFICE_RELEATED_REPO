package org.bitbudget.vegetable.office.poi.style.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
@Documented
public @interface ReportInfoVo {

	
	/**
	 * ��������Ӧ���к�
	 * "A"��"b"��"A-B" 
	 */
	public abstract String colNums();
	
	/**
	 * ���뷽ʽ
	 * 
	 */
	public abstract AlignmentType alignment() default AlignmentType.center;
	
	/**
	 *	�Ƿ�Ӵ� 
	 */
	public abstract boolean isBold() default false;
	
	/**
	 * �����С 
	 */
	public abstract String fontSize() default "14";
	
	/**
	 * ����ɫ 
	 */
	public abstract short backColor() default HSSFColor.WHITE.index;
	
	/**
	 * ������ɫ 
	 */
	public abstract short fontColor() default HSSFColor.BLACK.index;
	
}
