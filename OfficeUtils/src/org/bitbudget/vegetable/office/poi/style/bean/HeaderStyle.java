package org.bitbudget.vegetable.office.poi.style.bean;

import org.apache.poi.hssf.util.HSSFColor;

public class HeaderStyle {

	private String title;
	private String rowNums;
	private String colNums;
	private short backColor = HSSFColor.WHITE.index;
	private short fontColor = HSSFColor.BLACK.index;
	
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getRowNums() {
		return rowNums;
	}
	public void setRowNums(String rowNums) {
		this.rowNums = rowNums;
	}
	public String getColNums() {
		return colNums;
	}
	public void setColNums(String colNums) {
		this.colNums = colNums;
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
	
}
