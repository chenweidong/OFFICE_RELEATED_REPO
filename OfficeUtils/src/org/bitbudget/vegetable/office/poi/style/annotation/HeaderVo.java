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
	 * 导出excel的标题
	 */
	public abstract String title();

	/**
	 * 标题所在行号
	 * "1"或"2"或"1-2"
	 */
	public abstract String rowNums();
	
	/**
	 * 标题所对应的列号
	 * "A"或"b"或"A-B" 
	 */
	public abstract String colNums();
	
	/**
	 * 背景色 
	 */
	public abstract short backColor() default HSSFColor.WHITE.index;
	
	/**
	 * 字体颜色 
	 */
	public abstract short fontColor() default HSSFColor.BLACK.index;
	
	
	
}
