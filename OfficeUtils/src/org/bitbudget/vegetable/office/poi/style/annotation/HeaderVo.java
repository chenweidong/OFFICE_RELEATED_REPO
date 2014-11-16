package org.bitbudget.vegetable.office.poi.style.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.hssf.util.HSSFColor;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
@Documented
public @interface HeaderVo {

	/**
	 * ����excel�ı���
	 */
	public abstract String title();

	/**
	 * ���������к�
	 * "1"��"2"��"1-2"
	 */
	public abstract String rowNums();
	
	/**
	 * ��������Ӧ���к�
	 * "A"��"b"��"A-B" 
	 */
	public abstract String colNums();
	
	/**
	 * ����ɫ 
	 */
	public abstract short backColor() default HSSFColor.WHITE.index;
	
	/**
	 * ������ɫ 
	 */
	public abstract short fontColor() default HSSFColor.BLACK.index;
	
	
	
}
