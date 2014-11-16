package org.bitbudget.vegetable.office.poi.style.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD,ElementType.TYPE})
@Documented
public @interface BodyVo {

	/**
	 *	�к�,ָ���ֶ��������뵽excel����Ӧ��λ��
	 *	"A",��"A-c" 
	 */
	public abstract String colNums() default "A";
	
	/**
	 * �Ƿ����
	 */
	public abstract boolean rowSpan() default false;
	
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
	 * ����ɫ 
	 */
	public abstract short backColor() default HSSFColor.WHITE.index;
	
	/**
	 * ������ɫ 
	 */
	public abstract short fontColor() default HSSFColor.BLACK.index;
	
	
	/**
	 * ָ�����������������б�����һ�����ݣ��ֿ�������У�
	 */
	public abstract LayoutType contentType() default LayoutType.block;
	
}
