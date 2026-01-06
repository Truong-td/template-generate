package com.truongtd.templategenerate.service.impl;

import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.helper.TemplateContextBuilder;
import com.truongtd.templategenerate.request.GenerateTemplateRequest;
import com.truongtd.templategenerate.service.DraftingFlexDataService;
import com.truongtd.templategenerate.service.TableDataService;
import com.truongtd.templategenerate.service.TemplateService;
import com.truongtd.templategenerate.service.TextDataService;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;

import static com.truongtd.templategenerate.util.StringUtils.COND_END;
import static com.truongtd.templategenerate.util.StringUtils.COND_START;

@Slf4j
@Service
public class TemplateServiceImpl implements TemplateService {

    private final DocxTemplateEngine templateEngine = new DocxTemplateEngine();
    private final TemplateContextBuilder contextBuilder = new TemplateContextBuilder();
    private final TableDataService tableDataService;
    private final TextDataService textDataService;


    private final DraftingFlexDataService draftingFlexDataService;

//    private final MinIOService minIOService;

    public TemplateServiceImpl(TableDataService tableDataService, TextDataService textDataService,
                               DraftingFlexDataService draftingFlexDataService) {
        this.tableDataService = tableDataService;
        this.textDataService = textDataService;
        this.draftingFlexDataService = draftingFlexDataService;
    }
    @Override
    public byte[] generateDocx(GenerateTemplateRequest request) throws Exception {
        if (StringUtils.isEmpty(request.getTemplateCode())) {
            log.error("Error: no document found with code: {}", request.getTemplateCode());
        }
        try {
//            WordprocessingMLPackage pkg = loadTemplate(request.getTemplateCode());
            WordprocessingMLPackage pkg = WordprocessingMLPackage.load(
                    getClass().getResourceAsStream("/templates/template-report.docx"));

            Map<String, Object> context = contextBuilder.buildRootContext(request);

            templateEngine.fixLocationDateLayout(pkg);
            // 1. FlexData: thay {{key}} bằng text / table / image
            draftingFlexDataService.processFlexData(pkg, request.getFlexData());

            // 1) condition trước để xóa được cả table
            processConditionalBlocks(pkg, context);

            // 2. TableData: lặp các bảng list + custom "Danh sách môn học"
            tableDataService.processTableData(pkg, context);

            // 3. TextData: scalar + block {{?key}}...{{/key}}
            textDataService.processTextBlocks(pkg, context);

            // 4. Dọn các paragraph rỗng dư thừa
            cleanupEmptyParagraphs(pkg);

            // 6. Xuất docx ra mảng byte
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            pkg.save(out);
            return out.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Error generating template", e);
        }
    }

    private void processConditionalBlocks(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {
        Body body = pkg.getMainDocumentPart().getContents().getBody();
        List<Object> c = body.getContent();

        int i = 0;
        while (i < c.size()) {
            Object u = XmlUtils.unwrap(c.get(i));
            if (!(u instanceof P)) { i++; continue; }

            String s = Optional.ofNullable(templateEngine.getParagraphText((P) u)).orElse("").trim();
            Matcher ms = COND_START.matcher(s);
            if (!ms.matches()) { i++; continue; }

            String key = ms.group(1).trim().replaceAll("\\s+", ""); // remove spaces in key

            int end = -1;
            for (int j = i + 1; j < c.size(); j++) {
                Object uj = XmlUtils.unwrap(c.get(j));
                if (uj instanceof P) {
                    String ej = Optional.ofNullable(templateEngine.getParagraphText((P) uj)).orElse("").trim();
                    Matcher me = COND_END.matcher(ej);
                    if (me.matches()) {
                        String endKey = me.group(1).trim().replaceAll("\\s+", "");
                        if (endKey.equals(key)) { end = j; break; }
                    }
                }
            }
            if (end == -1) { i++; continue; }

            Object condVal = templateEngine.resolveKey(root, key);
            boolean show = templateEngine.isTruthy(condVal);

            if (!show) {
                // FALSE: xoá toàn block (kể cả TABLE)
                c.subList(i, end + 1).clear();
                continue;
            }

            // TRUE: chỉ xoá marker nếu block chứa TABLE (để tránh ảnh hưởng user/text block)
            boolean hasTblInside = false;
            for (int k = i + 1; k < end; k++) {
                Object uk = XmlUtils.unwrap(c.get(k));
                if (uk instanceof Tbl) { hasTblInside = true; break; }
            }

            if (hasTblInside) {
                // block có TABLE: xoá marker start/end, giữ nội dung
                c.remove(end); // end marker paragraph
                c.remove(i);   // start marker paragraph
                continue;
            }

            // block chỉ có text: giữ marker để processTextBlocks xử lý (merge context user)
            i++;
        }
    }

    private void cleanupEmptyParagraphs(WordprocessingMLPackage pkg) throws Docx4JException {
        Body body = pkg.getMainDocumentPart().getContents().getBody();
        List<Object> content = body.getContent();

        boolean prevEmpty = false;
        Iterator<Object> it = content.iterator();

        while (it.hasNext()) {
            Object o = XmlUtils.unwrap(it.next());
            if (!(o instanceof P)) {
                prevEmpty = false;
                continue;
            }
            P p = (P) o;
            String txt = templateEngine.getParagraphText(p).trim();
            boolean isEmpty = txt.isEmpty();

            if (isEmpty) {
                if (prevEmpty) {
                    it.remove(); // 2 paragraph trống liên tiếp -> bỏ cái sau
                } else {
                    prevEmpty = true;
                }
            } else {
                prevEmpty = false;
            }
        }
    }

//    private WordprocessingMLPackage loadTemplate(String templateCode) throws Docx4JException {
//
//        DocumentEntity approvalDocument = documentRepository.findFirstByCodeAndStatusOrderByFormTempIdDescVersionDesc(templateCode, DocumentStatusEnums.ACTIVE.name());
//        if (null == approvalDocument) {
//            log.error("Error: no document found with code: {}", templateCode);
//            throw new ApplicationException(DomainCode.INVALID_PARAMETER,
//                    new Object[]{"no document found with code " + templateCode});
//        }
//        log.info("Generate template from MinIO code: [{}], path: [{}]", templateCode,
//                approvalDocument.getPath());
//        InputStream fileDoc = minIOService.downloadFileMinIO(approvalDocument.getPath());
//        if (null == fileDoc) {
//            throw new IllegalStateException("Find not found template document from MinIO by code");
//        }
//        return WordprocessingMLPackage.load(fileDoc);
//    }
}
