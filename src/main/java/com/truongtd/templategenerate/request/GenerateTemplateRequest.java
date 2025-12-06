package com.truongtd.templategenerate.request;

import lombok.Data;

@Data
public class GenerateTemplateRequest {
    private String textData;

    private String tableData;

    private String flexData;

    public String getTextData() {
        return textData;
    }

    public void setTextData(String textData) {
        this.textData = textData;
    }

    public String getTableData() {
        return tableData;
    }

    public void setTableData(String tableData) {
        this.tableData = tableData;
    }

    public String getFlexData() {
        return flexData;
    }

    public void setFlexData(String flexData) {
        this.flexData = flexData;
    }
}
