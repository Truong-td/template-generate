package com.truongtd.templategenerate.service;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.util.Map;

public interface TableDataService {
    void processTableData(WordprocessingMLPackage pkg, Map<String, Object> root) ;
}
