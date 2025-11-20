package com.truongtd.templategenerate.dto;

import lombok.Data;

import java.util.List;

@Data
public class FlexDataDTO {

    private String image;

    private String text;

    private List<List<String>> table;
}
