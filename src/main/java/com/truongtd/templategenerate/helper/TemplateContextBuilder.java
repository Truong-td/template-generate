package com.truongtd.templategenerate.helper;

import com.truongtd.templategenerate.dto.TemplateDataDto;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TemplateContextBuilder {

    public Map<String, Object> buildContext(TemplateDataDto request) {
        Map<String, Object> root = new HashMap<>();

        if (request.getTextData() != null) {
            root.putAll(request.getTextData());
        }
        if (request.getTableData() != null) {
            root.putAll(request.getTableData());
        }
        if (request.getFlexData() != null) {
            root.putAll(request.getFlexData());
        }

        return root;
    }
}
