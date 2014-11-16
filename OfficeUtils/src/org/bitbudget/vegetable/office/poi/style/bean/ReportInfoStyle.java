package org.bitbudget.vegetable.office.poi.style.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;

public class ReportInfoStyle {

	private String colNums = "A";
	private AlignmentType alignment = AlignmentType.center;
	private boolean isBold = false;
	private short fontSize = 14;
	private short backColor = HSSFColor.WHITE.index;
	private short fontColor = HSSFColor.BLACK.index;
	private String text;
	
	public String getColNums() {
		return colNums;
	}
	public void setColNums(String colNums) {
		this.colNums = colNums;
	}
	public AlignmentType getAlignment() {
		return alignment;
	}
	public void setAlignment(AlignmentType alignment) {
		this.alignment = alignment;
	}
	public boolean getIsBold() {
		return isBold;
	}
	public void setBold(boolean isBold) {
		this.isBold = isBold;
	}
	public short getFontSize() {
		return fontSize;
	}
	public void setFontSize(short fontSize) {
		this.fontSize = fontSize;
	}
	public short getBackColor() {
		return backColor;
	}
	public void setBackColor(short backColor) {
		this.backColor = backColor;
	}
	public short getFontColor() {
		return fontColor;
	}
	public void setFontColor(short fontColor) {
		this.fontColor = fontColor;
	}
	public String getText() {
		return text;
	}
	public void setText(String text) {
		this.text = text;
	}
}
