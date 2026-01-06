package com.truongtd.templategenerate.service;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.util.Map;

public interface DraftingFlexDataService {

    void processFlexData(WordprocessingMLPackage pkg, Map<String, Object> flexData) throws Exception;
}
