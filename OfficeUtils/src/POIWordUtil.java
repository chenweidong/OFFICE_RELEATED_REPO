import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

public class POIWordUtil {

    public static void main(String[] args) throws Exception {
        Map<String, String> replaces = new HashMap<String, String>();
 
        replaces.put("${username}", "rongzhi_li");
        replaces.put("${password}", "1123456");
        replaces.put("${author}", "lee");

//        poiWordTableReplace("c:\\itext\\test.doc", "c:\\itext\\test2.doc", replaces);
       /* Exception in thread "main" java.lang.IllegalArgumentException: The document is really a RTF file
    	at org.apache.poi.hwpf.HWPFDocumentCore.verifyAndBuildPOIFS(HWPFDocumentCore.java:100)
    	at org.apache.poi.hwpf.HWPFDocument.<init>(HWPFDocument.java:174)
    	at POIWordUtil.poiWordTableReplace(POIWordUtil.java:30)
    	at POIWordUtil.main(POIWordUtil.java:24)*/
        
//        poiWordTableReplace("c:\\itext\\1.docx", "c:\\itext\\test2.docx", replaces);
        /*Exception in thread "main" org.apache.poi.poifs.filesystem.OfficeXmlFileException: The supplied data appears to be in the Office 2007+ XML. You are calling the part of POI that deals with OLE2 Office Documents. You need to call a different part of POI to process this data (eg XSSF instead of HSSF)
    	at org.apache.poi.poifs.storage.HeaderBlock.<init>(HeaderBlock.java:131)
    	at org.apache.poi.poifs.storage.HeaderBlock.<init>(HeaderBlock.java:104)
    	at org.apache.poi.poifs.filesystem.POIFSFileSystem.<init>(POIFSFileSystem.java:138)
    	at org.apache.poi.hwpf.HWPFDocumentCore.verifyAndBuildPOIFS(HWPFDocumentCore.java:106)
    	at org.apache.poi.hwpf.HWPFDocument.<init>(HWPFDocument.java:174)
    	at POIWordUtil.poiWordTableReplace(POIWordUtil.java:38)
    	at POIWordUtil.main(POIWordUtil.java:31)*/
        
//        poiWordTableReplace("c:\\itext\\1.doc", "c:\\itext\\test2.doc", replaces);
        ///打不开

    }

    public static void poiWordTableReplace(String sourceFile, String newFile,
            Map<String, String> replaces) throws Exception {
        FileInputStream in = new FileInputStream(sourceFile);
        HWPFDocument hwpf = new HWPFDocument(in);
        Range range = hwpf.getRange();// 得到文档的读取范围
        TableIterator it = new TableIterator(range);
        // 迭代文档中的表格
        while (it.hasNext()) {
            Table tb = (Table) it.next();
            // 迭代行，默认从0开始
            for (int i = 0; i < tb.numRows(); i++) {
                TableRow tr = tb.getRow(i);
                // 迭代列，默认从0开始
                for (int j = 0; j < tr.numCells(); j++) {
                    TableCell td = tr.getCell(j);// 取得单元格
                    // 取得单元格的内容
                    for (int k = 0; k < td.numParagraphs(); k++) {
                        Paragraph para = td.getParagraph(k);

                        String s = para.text();
                        final String old = s;
                        for (String key : replaces.keySet()) {
                            if (s.contains(key)) {
                                s = s.replace(key, replaces.get(key));
                            }
                        }
                        if (!old.equals(s)) {// 有变化
                            para.replaceText(old, s);
                            s = para.text();
                            System.out.println("old:" + old + "->" + "s:" + s);
                        }

                    } // end for
                } // end for
            } // end for
        } // end while

        FileOutputStream out = new FileOutputStream(newFile);
        hwpf.write(out);

        out.flush();
        out.close();

    }
//    public abstract class Text{
//    	
//    	public abstract String getText();
//    	
//    	public static Text str(final String string) {
//    		return new Text() {
//    			@Override
//    			public String getText() {
//    				return string;
//    			}
//    		};
//    	}
//    }
}