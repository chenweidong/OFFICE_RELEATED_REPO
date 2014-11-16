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
	 * 标题所对应的列号
	 * "A"或"b"或"A-B" 
	 */
	public abstract String colNums();
	
	/**
	 * 对齐方式
	 * 
	 */
	public abstract AlignmentType alignment() default AlignmentType.center;
	
	/**
	 *	是否加粗 
	 */
	public abstract boolean isBold() default false;
	
	/**
	 * 字体大小 
	 */
	public abstract String fontSize() default "14";
	
	/**
	 * 背景色 
	 */
	public abstract short backColor() default HSSFColor.WHITE.index;
	
	/**
	 * 字体颜色 
	 */
	public abstract short fontColor() default HSSFColor.BLACK.index;
	
}
