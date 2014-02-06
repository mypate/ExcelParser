package com.pate.excel;

public class CellData {
	public static final int NUMBERIC_CELL = 0;
	public static final int STRING_CELL = 1;
	/**
	 * 顺序编号（当一个单元格的字符长度大于50个时用多个CellData对象存储）
	 */
	private Integer sequence;
	public Integer rowIndex;
	private Integer cellIndex;
	private Integer dataType;
	private Double numberValue;
	private String stringValue;
	private Object value;
	
	public CellData() {
		sequence = 0;
	}
	
	public Integer getSequence() {
		return sequence;
	}
	public void setSequence(Integer sequence) {
		this.sequence = sequence;
	}
	public Integer getRowIndex() {
		return rowIndex;
	}
	public void setRowIndex(Integer rowIndex) {
		this.rowIndex = rowIndex;
	}
	public Integer getCellIndex() {
		return cellIndex;
	}
	public void setCellIndex(Integer cellIndex) {
		this.cellIndex = cellIndex;
	}
	public Integer getDataType() {
		return dataType;
	}
	public void setDataType(Integer dataType) {
		this.dataType = dataType;
	}
	public Double getNumberValue() {
		return numberValue;
	}
	public void setNumberValue(Double numberValue) {
		this.numberValue = numberValue;
	}
	public String getStringValue() {
		return stringValue;
	}
	public void setStringValue(String stringValue) {
		this.stringValue = stringValue;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}
	
}
