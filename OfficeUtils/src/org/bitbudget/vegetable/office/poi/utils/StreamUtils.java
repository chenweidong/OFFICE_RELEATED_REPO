package org.bitbudget.vegetable.office.poi.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class StreamUtils {

	private static StreamUtils instance = null;
	
	private StreamUtils(){};
	
	public static StreamUtils getStreamUtils(){
		if(instance == null){
			instance = new StreamUtils();
		}
		
		return instance;
	}
	
	public InputStream hssfWorkBook2InputStream(HSSFWorkbook workbook) throws IOException{
			
			if(workbook == null){
				return null;
			}
		 	ByteArrayOutputStream baos = new ByteArrayOutputStream();
	        workbook.write(baos);
	        baos.flush();
	        byte[] aa = baos.toByteArray();
	        InputStream excelStream = new ByteArrayInputStream(aa, 0, aa.length);
	        baos.close();
	        
	        return excelStream;
	        
	 }
	
	public InputStream xwpfDocument2InputStream(XWPFDocument document) throws IOException{
        
		if(document == null){
			return null;
		}
	 	ByteArrayOutputStream baos = new ByteArrayOutputStream();
	 	document.write(baos);
        baos.flush();
        byte[] aa = baos.toByteArray();
        InputStream excelStream = new ByteArrayInputStream(aa, 0, aa.length);
        baos.close();
        
        return excelStream;
	}
}
