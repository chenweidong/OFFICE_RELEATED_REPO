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
	 *	列号,指定字段内容输入到excel的相应列位置
	 *	"A",或"A-c" 
	 */
	public abstract String colNums() default "A";
	
	/**
	 * 是否跨行
	 */
	public abstract boolean rowSpan() default false;
	
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
	 * 背景色 
	 */
	public abstract short backColor() default HSSFColor.WHITE.index;
	
	/**
	 * 字体颜色 
	 */
	public abstract short fontColor() default HSSFColor.BLACK.index;
	
	
	/**
	 * 指定数据内容是整行列表，还是一行数据（分块或者整行）
	 */
	public abstract LayoutType contentType() default LayoutType.block;
	
}
