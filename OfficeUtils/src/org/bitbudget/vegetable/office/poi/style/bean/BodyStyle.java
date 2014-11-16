package org.bitbudget.vegetable.office.poi.style.bean;

import org.apache.poi.hssf.util.HSSFColor;
import org.bitbudget.vegetable.office.poi.constant.AlignmentType;
import org.bitbudget.vegetable.office.poi.constant.LayoutType;

public class BodyStyle {

	private String colNums = "0";
	private boolean rowSpan = false;
	private AlignmentType alignment = AlignmentType.center;
	private boolean isBold = false;
	private short backColor = HSSFColor.WHITE.index;
	private short fontColor = HSSFColor.BLACK.index;
	private LayoutType contentType = LayoutType.block;
	private String text;
	
	public String getColNums() {
		return colNums;
	}
	public void setColNums(String colNums) {
		this.colNums = colNums;
	}
	public boolean getRowSpan() {
		return rowSpan;
	}
	public void setRowSpan(boolean rowSpan) {
		this.rowSpan = rowSpan;
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
	public LayoutType getContentType() {
		return contentType;
	}
	public void setContentType(LayoutType contentType) {
		this.contentType = contentType;
	}
	public String getText() {
		return text;
	}
	public void setText(String text) {
		this.text = text;
	}
	
	
}
