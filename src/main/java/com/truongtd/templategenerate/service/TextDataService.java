package com.truongtd.templategenerate.service;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.util.Map;

public interface TextDataService {
    void processTextBlocks(WordprocessingMLPackage pkg, Map<String, Object> context) throws Docx4JException;
}
