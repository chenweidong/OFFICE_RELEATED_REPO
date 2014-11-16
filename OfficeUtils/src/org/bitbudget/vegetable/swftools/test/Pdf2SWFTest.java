package org.bitbudget.vegetable.swftools.test;

import java.io.IOException;

import junit.framework.TestCase;

public class Pdf2SWFTest extends TestCase{

	public void testPdf2SWF() throws IOException{
		
        String command = "C:\\Program Files\\SWFTools\\pdf2swf.exe" + " -f -o \"" + 
        					"c:\\test\\" + "%.swf\" -s languagedir=" + "D:\\xpdf\\xpdf-chinese-simplified" + 
        					" -s flashversion=9 \"" + "D:\\Backup\\ÎÒµÄÎÄµµ\\e2.pdf" + "\"";
        
        Runtime.getRuntime().exec(command);
	}
}
