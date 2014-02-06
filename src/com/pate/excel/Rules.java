package com.pate.excel;

import java.util.ArrayList;
import java.util.List;

import com.google.gson.Gson;

public class Rules {
    public static List<ObjectRule> getRules() throws Exception { 
        Gson gson = new Gson();
        List<ObjectRule> objectRuleList = new ArrayList<ObjectRule>();
        ObjectRule or = null;
        or = gson.fromJson(getJson0(), ObjectRule.class);
        objectRuleList.add(or);
        initial(or);
        or = gson.fromJson(getJson1(), ObjectRule.class);
        initial(or);
        objectRuleList.add(or);
        return objectRuleList;
    }
    
    private static void initial(ObjectRule rule) throws Exception {
        List<FieldRule> list = rule.getFieldRuleList();
        String fieldName;
        @SuppressWarnings("unchecked")
		Class<Object> objectClass = (Class<Object>) Class.forName(rule.getObjectName());
        rule.setObjectClass(objectClass);
        for (FieldRule fieldRule : list) {
            int startRow = fieldRule.getStartRow();
            int endRow = fieldRule.getEndRow();
            int startCol = fieldRule.getStartCol();
            int endCol = fieldRule.getEndCol();
            if (startRow == endRow) {
                if (startCol != endCol) {
                    fieldRule.setPositionType(FieldRule.COLUMN_POSITION);
                } else {
                    fieldRule.setPositionType(FieldRule.NONE);
                }
            } else {
                fieldRule.setPositionType(FieldRule.ROW_POSITION);;
            }
            
            fieldName = fieldRule.getFieldName();
            if (fieldName == null || fieldName.trim().isEmpty()) continue;
            fieldRule.setField(objectClass.getField(fieldName));
        }
    }
    
    private static String getJson0() {
        return "{\"objectName\":\"com.pate.excel.PriceTable\","
                + "fieldRuleList:["
                + "{\"fieldName\":\"buildingNumber\","
                + "\"startCol\":1,"
                + "\"endCol\":1,"
                + "\"startRow\":1,"
                + "\"endRow\":1},"
                + "{\"fieldName\":\"roomNumber\","
                + "\"startCol\":0,"
                + "\"endCol\":0,"
                + "\"startRow\":5,"
                + "\"endRow\":23},"
                + "{\"fieldName\":\"price\","
                + "\"startCol\":2,"
                + "\"endCol\":2,"
                + "\"startRow\":5,"
                + "\"endRow\":23}" +
                "]}";
    }

    private static String getJson1() {
        return "{\"objectName\":\"com.pate.excel.PriceTable\","
                + "fieldRuleList:["
                + "{\"fieldName\":\"buildingNumber\","
                + "\"startCol\":3,"
                + "\"endCol\":3,"
                + "\"startRow\":1,"
                + "\"endRow\":1},"
                + "{\"fieldName\":\"roomNumber\","
                + "\"startCol\":0,"
                + "\"endCol\":0,"
                + "\"startRow\":5,"
                + "\"endRow\":23},"
                + "{\"fieldName\":\"price\","
                + "\"startCol\":4,"
                + "\"endCol\":4,"
                + "\"startRow\":5,"
                + "\"endRow\":23}" +
                "]}";
    }    
}
