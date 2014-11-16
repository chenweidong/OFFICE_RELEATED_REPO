package org.bitbudget.vegetable.test.annotation;

import org.bitbudget.vegetable.office.poi.test.ExcelVoTest;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;

import junit.framework.TestCase;

public class AnnotationTest extends TestCase{

	/**
	 * 测试如何获取类型上面声明的注解，见ExcelVoTest类名称上方的@BodyVo
	 * 	@BodyVo的contentType值为entire
	 */
	public void testGetClassTypeAnnotation(){
//		Class<?> declaringClass = ExcelVoTest.class.getDeclaringClass();
//		if(declaringClass != null){
//			if(declaringClass.isAnnotationPresent(BodyVo.class)){
//				System.out.println(declaringClass.getAnnotation(BodyVo.class).contentType());
//			}
//		}else{
//			System.out.println(declaringClass != null);
//		}
		
		//////result/////// 
		//false
		
//////////////////////////////////////////////////////////////////////////////////				
/////////////////////////////////////////////////////////////////////////////////
		
//		Class<?>[] declaredClasses = ExcelVoTest.class.getDeclaredClasses();
//		if(declaredClasses != null){
//			if(declaredClasses.length > 0){
//				if(declaredClasses[0].isAnnotationPresent(BodyVo.class)){
//					System.out.println(declaredClasses[0].getAnnotation(BodyVo.class).contentType());
//				}
//			}else{
//				System.out.println("declaredClasses.length = " + declaredClasses.length);
//			}
//		}else{
//			System.out.println(declaredClasses != null);
//		}
		
		//////result/////// 
		//declaredClasses.length = 0
		
		
//////////////////////////////////////////////////////////////////////////////////				
/////////////////////////////////////////////////////////////////////////////////
		
		if(ExcelVoTest.class.isAnnotationPresent(BodyVo.class)){
			System.out.println(ExcelVoTest.class.getAnnotation(BodyVo.class).contentType());
		}else{
			System.out.println("ExcelVoTest.class.isAnnotationPresent(BodyVo.class) = " + ExcelVoTest.class.isAnnotationPresent(BodyVo.class));
		}
		
		//////result/////// 
		//entire
		
	}
}
