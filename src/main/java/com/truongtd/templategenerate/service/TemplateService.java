package com.truongtd.templategenerate.service;

import com.truongtd.templategenerate.request.CreateTemplateRequest;

public interface TemplateService {
    byte[] generateDocx(CreateTemplateRequest request) throws Exception;
}
