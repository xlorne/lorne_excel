package com.lorne.core.framework.utils.excel.model;


/**
 * 封装合并单元格中的数据
 * @author yuliang
 *
 */
public class LMergeCellModel {

	private String value;
	private int startY;
	private int endY;
	
	private int startX;
	private int endX;
	
	
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	public int getStartY() {
		return startY;
	}
	public void setStartY(int startY) {
		this.startY = startY;
	}
	public int getEndY() {
		return endY;
	}
	public void setEndY(int endY) {
		this.endY = endY;
	}
	public int getStartX() {
		return startX;
	}
	public void setStartX(int startX) {
		this.startX = startX;
	}
	public int getEndX() {
		return endX;
	}
	public void setEndX(int endX) {
		this.endX = endX;
	}
	
	public LMergeCellModel() {
		
	}
	public LMergeCellModel(String value, int startY, int endY, int startX,
						   int endX) {
		super();
		this.value = value;
		this.startY = startY;
		this.endY = endY;
		this.startX = startX;
		this.endX = endX;
	}
	
	
	
	
}
