package com.pate.excel;

import java.util.List;

public class ObjectRule {
	/**
	 * �������Թ��򼯺�
	 */
	private List<FieldRule> fieldRuleList;
	
	private Class<Object> objectClass;
	
	/**
	 * ������������·����������
	 */
	private String objectName;

	public List<FieldRule> getFieldRuleList() {
		return fieldRuleList;
	}
	
	public Class<Object> getObjectClass() {
		return objectClass;
	}

	public String getObjectName() {
		return objectName;
	}

	public void setFieldRuleList(List<FieldRule> fieldRuleList) {
		this.fieldRuleList = fieldRuleList;
	}

	public void setObjectClass(Class<Object> objectClass) {
		this.objectClass = objectClass;
	}

	public void setObjectName(String objectName) {
		this.objectName = objectName;
	}
	
}
