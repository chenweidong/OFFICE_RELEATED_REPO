package org.bitbudget.vegetable.test.annotation;

import org.bitbudget.vegetable.office.poi.test.ExcelVoTest;
import org.bitbudget.vegetable.office.poi.style.annotation.BodyVo;

import junit.framework.TestCase;

public class AnnotationTest extends TestCase{

	/**
	 * ������λ�ȡ��������������ע�⣬��ExcelVoTest�������Ϸ���@BodyVo
	 * 	@BodyVo��contentTypeֵΪentire
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
