package com.truongtd.templategenerate.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.truongtd.templategenerate.dto.TemplateDataDto;
import com.truongtd.templategenerate.request.GenerateTemplateRequest;

import java.io.IOException;
import java.util.Collections;
import java.util.Map;

public class JsonUtils {
    private static final ObjectMapper MAPPER = new ObjectMapper();

    @SuppressWarnings("unchecked")
    public static Map<String, Object> parseToMap(String json) {
        if (json == null || json.trim().isEmpty()) return Collections.emptyMap();
        try {
            return MAPPER.readValue(json, new TypeReference<Map<String, Object>>() {});
        } catch (IOException e) {
            throw new RuntimeException("Cannot parse json", e);
        }
    }

    public static TemplateDataDto parse(GenerateTemplateRequest req) {
        TemplateDataDto data = new TemplateDataDto();
        data.setTextData(parseToMap(req.getTextData()));
        data.setTableData(parseToMap(req.getTableData()));
        data.setFlexData(parseToMap(req.getFlexData()));
        return data;
    }
}
