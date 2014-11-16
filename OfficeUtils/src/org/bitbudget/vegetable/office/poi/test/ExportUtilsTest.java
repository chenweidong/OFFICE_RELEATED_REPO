package org.bitbudget.vegetable.office.poi.test;

import org.bitbudget.vegetable.office.poi.utils.history.ExportUtils;

import junit.framework.TestCase;

public class ExportUtilsTest extends TestCase{

	public void testExport(){
		ExportUtils utils = new ExportUtils("ÏîÄ¿excel»ã×Ü");
		
		utils.exportExcel(ExcelVoTest.class);
	}
}
