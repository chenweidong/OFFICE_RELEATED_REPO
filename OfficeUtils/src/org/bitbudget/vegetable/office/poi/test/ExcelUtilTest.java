package org.bitbudget.vegetable.office.poi.test;

import junit.framework.Assert;

import org.bitbudget.vegetable.office.poi.utils.ExcelUtil;
import org.junit.Test;


public class ExcelUtilTest {

	/*
	 * A   B   C   D ...   Z
	 * 
	 * 26
	 * --------------------------
	 * 
	 * AA  AB  AC  AD      AZ
	 * BA 	...			   BZ
	 * ...
	 * ZA   ZB  ZC ...     ZZ
	 * 
	 * 26 + 26*26 = 702
	 * --------------------------
	 * 
	 * 
	 * AAA AAB  ...        AAZ
	 * ABA ABB  ABC   ...  ABZ
	 * ACA ACB  ACC   ...  ACZ
	 * ..
	 * BAA BAB  BAC   ...  BAZ
	 * BBA BBB  BBC   ...  BBZ
	 * ...
	 * ...
	 * ZAA ZAB ZAC   ...   ZAZ
	 * ZBA ZBB ZBC   ...   ZBZ
	 * ...
	 * ZZA ZZB ZZC   ...   ZZZ
	 * 
	 * 26 + 26*26 + 26*26*26 =18278
	 * --------------------------
	 */
	@Test
	public void testGetColAliaseByNum(){
		ExcelUtil instance = ExcelUtil.getInstance();
		
		Assert.assertEquals("A", instance.getColAliaseByNum(0));
		Assert.assertEquals("Z", instance.getColAliaseByNum(25));
		Assert.assertEquals("AA", instance.getColAliaseByNum(26));
		Assert.assertEquals("AZ", instance.getColAliaseByNum(51));
		Assert.assertEquals("BA", instance.getColAliaseByNum(52));
		Assert.assertEquals("BZ", instance.getColAliaseByNum(77));
		Assert.assertEquals("ZZ", instance.getColAliaseByNum(701));
//		Assert.assertEquals("AAA", instance.getColAliaseByNum(702));
//		Assert.assertEquals("ZZZ", instance.getColAliaseByNum(18277));
	}
}
