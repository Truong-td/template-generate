package com.truongtd.templategenerate.request;

import lombok.Data;

@Data
public class GenerateTemplateRequest {
    private String textData;

    private String tableData;

    private String flexData;
}
