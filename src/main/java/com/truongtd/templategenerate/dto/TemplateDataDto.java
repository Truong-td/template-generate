package com.truongtd.templategenerate.dto;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.util.Map;

@Data
public class TemplateDataDto {

    private Map<String, Object> textData;
    private Map<String, Object> tableData;
    private Map<String, Object> flexData;
}
