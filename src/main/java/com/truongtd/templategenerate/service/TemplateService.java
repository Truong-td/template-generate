package com.truongtd.templategenerate.service;

import com.truongtd.templategenerate.request.CreateTemplateRequest;
import com.truongtd.templategenerate.request.GenerateTemplateRequest;

public interface TemplateService {

    byte[] generateDocx(GenerateTemplateRequest request);
}
