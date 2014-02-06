package com.pate.excel;

import java.lang.reflect.Field;

public class FieldRule {
	public static final int COLUMN_POSITION = 1;
	public static final int ROW_POSITION = 2;
	public static final int NONE = 0;
    /**
     * �ֶ�����
     */
    private String fieldName;

    /**
     * ��λ�ã���ʼ��
     */
    private Integer startCol;

    /**
     * ��λ�ã�������
     */
    private Integer endCol;

    /**
     * ��λ�ã���ʼ��
     */
    private Integer startRow;

    /**
     * ��λ�ã�������
     */
    private Integer endRow;
    
    private Integer positionType;
    
    /*
     * �ֶ�setter����
     */
    private Field field;

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public Integer getStartCol() {
        return startCol;
    }

    public void setStartCol(Integer startCol) {
        this.startCol = startCol;
    }

    public Integer getEndCol() {
        return endCol;
    }

    public void setEndCol(Integer endCol) {
        this.endCol = endCol;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getEndRow() {
        return endRow;
    }

    public void setEndRow(Integer endRow) {
        this.endRow = endRow;
    }

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

	public Integer getPositionType() {
		return positionType;
	}

	public void setPositionType(Integer positionType) {
		this.positionType = positionType;
	}

}
