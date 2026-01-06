package com.truongtd.templategenerate.request;

import lombok.Data;

import java.util.Map;

@Data
public class GenerateTemplateRequest {
    private Map<String, Object> textData;

    private Map<String, Object> tableData;

    private Map<String, Object> flexData;

    private String templateCode;

    public Map<String, Object> getTextData() {
        return textData;
    }


    private String fileName;
    public void setTextData(Map<String, Object> textData) {
        this.textData = textData;
    }

    public Map<String, Object> getTableData() {
        return tableData;
    }

    public void setTableData(Map<String, Object> tableData) {
        this.tableData = tableData;
    }

    public Map<String, Object> getFlexData() {
        return flexData;
    }

    public void setFlexData(Map<String, Object> flexData) {
        this.flexData = flexData;
    }
}
