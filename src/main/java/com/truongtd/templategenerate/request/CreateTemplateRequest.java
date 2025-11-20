package com.truongtd.templategenerate.request;

import com.truongtd.templategenerate.dto.FlexDataDTO;
import lombok.Data;

import java.util.List;

@Data
public class CreateTemplateRequest {

    private String textData;

    private String tableData;

    private List<FlexDataDTO> flexDataList;
}
