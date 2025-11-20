package com.truongtd.templategenerate.mapper;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.truongtd.templategenerate.dto.FlexDataDTO;
import com.truongtd.templategenerate.dto.TemplateDataDto;
import com.truongtd.templategenerate.request.CreateTemplateRequest;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TemplateMapper {
    private static final ObjectMapper objectMapper = new ObjectMapper();

    public static TemplateDataDto convert(CreateTemplateRequest req) {

        TemplateDataDto dto = new TemplateDataDto();

        try {
            // textData: JSON string → Map<String, Object>
            if (req.getTextData() != null) {
                dto.setTextData(
                        objectMapper.readValue(req.getTextData(), new TypeReference<Map<String, Object>>() {})
                );
            }

            // tableData: JSON string → Map<String, Object>
            if (req.getTableData() != null) {
                dto.setTableData(
                        objectMapper.readValue(req.getTableData(), new TypeReference<Map<String, Object>>() {})
                );
            }

            // flexData: List<FlexDataDTO> → List<Map<String, Object>>
            if (req.getFlexDataList() != null) {
                List<Map<String, Object>> flexList = new ArrayList<>();

                for (FlexDataDTO item : req.getFlexDataList()) {
                    Map<String, Object> map = new HashMap<>();
                    map.put("image", item.getImage());
                    map.put("text", item.getText());
                    map.put("table", item.getTable());
                    flexList.add(map);
                }

                // Đưa vào flexData
                Map<String, Object> flexData = new HashMap<>();
                flexData.put("flexList", flexList);

                dto.setFlexData(flexData);
            }

        } catch (Exception e) {
            throw new RuntimeException("Convert request to TemplateDataDto failed", e);
        }

        return dto;
    }
}
