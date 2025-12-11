package com.truongtd.templategenerate.controller;

import com.truongtd.templategenerate.request.CreateTemplateRequest;
import com.truongtd.templategenerate.request.GenerateTemplateRequest;
import com.truongtd.templategenerate.service.TemplateService;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/templates")
public class GenerateTemplateController {

    private final TemplateService templateService;

    public GenerateTemplateController(TemplateService templateService) {
        this.templateService = templateService;
    }

    @PostMapping(
            value = "/generate/template-document",
            produces = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    public ResponseEntity<byte[]> generate(@RequestBody GenerateTemplateRequest request) throws Exception {
        byte[] fileBytes = templateService.generateDocx(request);

        String filename = "report-" + System.currentTimeMillis() + ".docx";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(
                MediaType.parseMediaType(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        );
        ContentDisposition contentDisposition = ContentDisposition
                .attachment()
                .filename(filename)
                .build();
        headers.setContentDisposition(contentDisposition);

        return new ResponseEntity<>(fileBytes, headers, HttpStatus.OK);
    }
}
