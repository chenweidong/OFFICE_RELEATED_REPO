package org.bitbudget.vegetable.office.poi.test;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakClear;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import junit.framework.TestCase;

public class ExportWordTest extends TestCase {

	 public static void main(String[] args) throws Exception {
	        XWPFDocument doc = new XWPFDocument();
	        
	        //设置标题--业务意义上的标题
	        XWPFParagraph p0 = doc.createParagraph();
	        p0.setIndentationFirstLine(500);
//	        p0.setAlignment(ParagraphAlignment.CENTER);
	        XWPFRun run0 = p0.createRun();
	        run0.setBold(true);
	        run0.setText("XX年X月XXX月报导出表");
	        run0.setFontSize(14);
	        

	        XWPFParagraph p1 = doc.createParagraph();
	        XWPFRun r1 = p1.createRun();
	        r1.setBold(true);
	        r1.setText("一、信息化建设概况");
	        r1.setFontSize(12);
	        r1.setTextPosition(10);

	        XWPFParagraph p2 = doc.createParagraph();
	        XWPFRun run2 = p2.createRun();
	        run2.setText("    1.信息化建设项目");
	        run2.setTextPosition(15);
	        
	        XWPFParagraph p3 = doc.createParagraph();
	        //BORDERS
	        p3.setBorderBottom(Borders.SINGLE);
	        p3.setBorderTop(Borders.SINGLE);
	        p3.setBorderRight(Borders.SINGLE);
	        p3.setBorderLeft(Borders.SINGLE);
	        p3.setBorderBetween(Borders.SINGLE);
	        p3.setAlignment(ParagraphAlignment.NUM_TAB);
	        XWPFRun r3 = p3.createRun();
	        r3.setText("jumped over the lazy dog 中国人\r\n\t民大团结");
	        
	        XWPFParagraph p22 = doc.createParagraph();
	        XWPFRun run22 = p22.createRun();
	        run22.setText("    2.信息基础设施建设");
	        run22.setTextPosition(15);
	        
	        XWPFParagraph p32 = doc.createParagraph();
	        XWPFRun r32 = p32.createRun();
	        r32.setText("jumped over the lazy dog 中国人 民大团结");
	        
	        
	        XWPFParagraph p4 = doc.createParagraph();
	        XWPFRun run4 = p4.createRun();
	        run4.setText("三、XX年X月工作完成情况表");
	        run4.setBold(true); 
	        run4.setFontSize(12);
	        run4.setTextPosition(10);
	        
	        XWPFTable table1 = doc.createTable(3,3);
	        table1.setWidth(50000);
	        table1.getRow(1).getCell(0).setText("试验");
	        table1.getRow(1).getCell(1).setText("test");
	        table1.setWidth(500);
////
//	        XWPFRun r3 = p2.createRun();
//	        r3.setText("and went away");
//	        r3.setStrike(true);
//	        r3.setFontSize(20);
//	        r3.setSubscript(VerticalAlign.SUPERSCRIPT);
//
//
//	        XWPFParagraph p3 = doc.createParagraph();
//	        p3.setWordWrap(true);
//	        p3.setPageBreak(true);
//	                
//	        //p3.setAlignment(ParagraphAlignment.DISTRIBUTE);
//	        p3.setAlignment(ParagraphAlignment.BOTH);
//	        p3.setSpacingLineRule(LineSpacingRule.EXACT);
//
//	        p3.setIndentationFirstLine(600);
//	        
//
//	        XWPFRun r4 = p3.createRun();
//	        r4.setTextPosition(20);
//	        r4.setText("To be, or not to be: that is the question: "
//	                + "Whether 'tis nobler in the mind to suffer "
//	                + "The slings and arrows of outrageous fortune, "
//	                + "Or to take arms against a sea of troubles, "
//	                + "And by opposing end them? To die: to sleep; ");
//	        r4.addBreak(BreakType.PAGE);
//	        r4.setText("No more; and by a sleep to say we end "
//	                + "The heart-ache and the thousand natural shocks "
//	                + "That flesh is heir to, 'tis a consummation "
//	                + "Devoutly to be wish'd. To die, to sleep; "
//	                + "To sleep: perchance to dream: ay, there's the rub; "
//	                + ".......");
//	        r4.setItalic(true);
//	//This would imply that this break shall be treated as a simple line break, and break the line after that word:
//
//	        XWPFRun r5 = p3.createRun();
//	        r5.setTextPosition(-10);
//	        r5.setText("For in that sleep of death what dreams may come");
//	        r5.addCarriageReturn();
//	        r5.setText("When we have shuffled off this mortal coil,"
//	                + "Must give us pause: there's the respect"
//	                + "That makes calamity of so long life;");
//	        r5.addBreak();
//	        r5.setText("For who would bear the whips and scorns of time,"
//	                + "The oppressor's wrong, the proud man's contumely,");
//	        
//	        r5.addBreak(BreakClear.ALL);
//	        r5.setText("The pangs of despised love, the law's delay,"
//	                + "The insolence of office and the spurns" + ".......");

	        FileOutputStream out = new FileOutputStream("src/org/bitbudget/vegetable/office/poi/test/simple.docx");
	        doc.write(out);
	        out.close();

	    }
}
