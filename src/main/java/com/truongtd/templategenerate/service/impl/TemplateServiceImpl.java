package com.truongtd.templategenerate.service.impl;

import com.truongtd.templategenerate.dto.TemplateDataDto;
import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.helper.TemplateContextBuilder;
import com.truongtd.templategenerate.mapper.TemplateMapper;
import com.truongtd.templategenerate.request.CreateTemplateRequest;
import com.truongtd.templategenerate.service.TemplateService;
import jakarta.annotation.PostConstruct;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Document;
import org.docx4j.wml.P;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.docx4j.wml.Text;
import org.docx4j.TraversalUtil.Callback;

@Service
public class TemplateServiceImpl implements TemplateService {

    private static final Logger log = LoggerFactory.getLogger(TemplateServiceImpl.class);

    private final DocxTemplateEngine templateEngine = new DocxTemplateEngine();
    private final TemplateContextBuilder contextBuilder = new TemplateContextBuilder();

    // {{blockName}}  &  {{/blockName}} (trên cả paragraph)
    private static final Pattern BLOCK_START_PATTERN =
            Pattern.compile("^\\s*\\{\\{(\\w+)}}\\s*$");
    private static final Pattern BLOCK_END_PATTERN =
            Pattern.compile("^\\s*\\{\\{/(\\w+)}}\\s*$");

    @PostConstruct
    public void checkTemplateExists() {
        try (InputStream is = getClass().getResourceAsStream("/templates/template-report.docx")) {
            if (is == null) {
                throw new IllegalStateException("Không tìm thấy /templates/template-report.docx trong resources");
            }
            log.info("Đã load được template-report.docx");
        } catch (IOException e) {
            throw new RuntimeException("Lỗi khi kiểm tra template-report.docx", e);
        }
    }

    @Override
    public byte[] generateDocx(CreateTemplateRequest request) throws Exception {
        TemplateDataDto templateDataDto = TemplateMapper.convert(request);

        Map<String, Object> rootContext = contextBuilder.buildContext(templateDataDto);
        log.debug("Context dùng để render: {}", rootContext);

        WordprocessingMLPackage wordMLPackage = loadTemplate();
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        processDocument(mainDocumentPart, rootContext);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wordMLPackage.save(out);
        return out.toByteArray();
    }

    private WordprocessingMLPackage loadTemplate() throws Docx4JException, IOException {
        try (InputStream is = getClass().getResourceAsStream("/templates/template-report.docx")) {
            if (is == null) {
                throw new IllegalStateException("Không tìm thấy template-report.docx trong /templates");
            }
            return WordprocessingMLPackage.load(is);
        }
    }

    /**
     * Xử lý toàn bộ document:
     * - Tìm block {{name}} ... {{/name}} theo thứ tự paragraph
     * - Nhân bản body block theo dữ liệu (object/list)
     * - Replace scalar {{key}} bên trong từng paragraph bằng DocxTemplateEngine
     */
    @SuppressWarnings("unchecked")
    private void processDocument(MainDocumentPart mainDocumentPart,
                                 Map<String, Object> rootContext) {

        Document wmlDoc = (Document) mainDocumentPart.getJaxbElement();
        List<Object> originalContent = wmlDoc.getBody().getContent();

        List<Object> newContent = new ArrayList<>();

        for (int i = 0; i < originalContent.size(); i++) {
            Object el = originalContent.get(i);
            Object unwrapped = XmlUtils.unwrap(el);

            if (!(unwrapped instanceof P)) {
                // Không phải paragraph => copy sang như cũ (table, hình, ...)
                newContent.add(XmlUtils.deepCopy(el));
                continue;
            }

            P p = (P) unwrapped;
            String paraText = getParagraphText(p);

            if (paraText == null) {
                newContent.add(XmlUtils.deepCopy(el));
                continue;
            }

            String trimmed = paraText.trim();

            // 1. Check START block {{name}}
            Matcher startMatcher = BLOCK_START_PATTERN.matcher(trimmed);
            if (startMatcher.matches()) {
                String blockName = startMatcher.group(1);
                log.debug("Gặp block START: {} tại index {}", blockName, i);

                // Tìm END {{/name}}
                int endIndex = -1;
                for (int j = i + 1; j < originalContent.size(); j++) {
                    Object elEnd = originalContent.get(j);
                    Object uEnd = XmlUtils.unwrap(elEnd);
                    if (uEnd instanceof P pEnd) {
                        String endText = getParagraphText(pEnd);
                        if (endText != null) {
                            Matcher endMatcher =
                                    BLOCK_END_PATTERN.matcher(endText.trim());
                            if (endMatcher.matches()
                                    && blockName.equals(endMatcher.group(1))) {
                                endIndex = j;
                                break;
                            }
                        }
                    }
                }

                if (endIndex == -1) {
                    // Không tìm thấy end -> coi như paragraph thường
                    log.warn("Không tìm thấy END cho block {}, xử lý như paragraph thường", blockName);
                    P cloned = (P) XmlUtils.deepCopy(p);
                    replaceScalarsInParagraph(cloned, rootContext);
                    newContent.add(cloned);
                    continue;
                }

                log.debug("Block {}: START {} - END {}", blockName, i, endIndex);

                // Body paragraphs: (i+1) .. (endIndex-1)
                List<Object> body = originalContent.subList(i + 1, endIndex);

                Object blockData = rootContext.get(blockName);

                if (blockData instanceof List<?> list) {
                    // Block lặp list
                    for (Object item : list) {
                        Map<String, Object> subContext = castToMap(item);
                        for (Object bodyEl : body) {
                            Object copy = XmlUtils.deepCopy(bodyEl);
                            Object bodyUnwrapped = XmlUtils.unwrap(copy);
                            if (bodyUnwrapped instanceof P bodyPara) {
                                replaceScalarsInParagraph(bodyPara, subContext);
                            }
                            newContent.add(copy);
                        }
                    }
                } else if (blockData instanceof Map<?, ?> map) {
                    // Block object
                    Map<String, Object> subContext = (Map<String, Object>) map;
                    for (Object bodyEl : body) {
                        Object copy = XmlUtils.deepCopy(bodyEl);
                        Object bodyUnwrapped = XmlUtils.unwrap(copy);
                        if (bodyUnwrapped instanceof P bodyPara) {
                            replaceScalarsInParagraph(bodyPara, subContext);
                        }
                        newContent.add(copy);
                    }
                } else {
                    // Không có data / không phải map/list: bỏ cả block
                    log.debug("Block {} không có dữ liệu phù hợp, bỏ qua toàn block", blockName);
                }

                // Nhảy qua đoạn đã xử lý (bao gồm END)
                i = endIndex;
                continue;
            }

            // 2. Check END block "mồ côi" -> bỏ
            Matcher endMatcher = BLOCK_END_PATTERN.matcher(trimmed);
            if (endMatcher.matches()) {
                log.warn("Gặp END block mồ côi: {} tại index {}, bỏ qua", endMatcher.group(1), i);
                continue;
            }

            // 3. Paragraph bình thường: replace scalar với rootContext
            P cloned = (P) XmlUtils.deepCopy(p);
            replaceScalarsInParagraph(cloned, rootContext);
            newContent.add(cloned);
        }

        // Ghi lại nội dung mới cho body
        originalContent.clear();
        originalContent.addAll(newContent);
    }

    /**
     * Lấy full text của 1 paragraph (gộp tất cả w:t bên trong)
     */
    private String getParagraphText(P paragraph) {
        List<Text> texts = new ArrayList<>();

        new TraversalUtil(paragraph, new Callback() {
            @Override
            public List<Object> apply(Object o) {
                Object u = XmlUtils.unwrap(o);
                if (u instanceof Text t) {
                    texts.add(t);
                }
                return null;
            }

            @Override
            public boolean shouldTraverse(Object o) {
                return true;
            }

            @Override
            public void walkJAXBElements(Object parent) {
                List<Object> children = getChildren(parent);
                if (children != null) {
                    for (Object o : children) {
                        Object u = XmlUtils.unwrap(o);
                        apply(u);
                        walkJAXBElements(u);
                    }
                }
            }

            @Override
            public List<Object> getChildren(Object o) {
                return TraversalUtil.getChildrenImpl(o);
            }
        });

        if (texts.isEmpty()) return null;

        StringBuilder sb = new StringBuilder();
        for (Text t : texts) {
            sb.append(t.getValue());
        }
        return sb.toString();
    }

    /**
     * Replace scalar {{key}} trong 1 paragraph với context cho trước.
     * Giữ nguyên cấu trúc w:r/w:t, chỉ chỉnh text.
     */
    private void replaceScalarsInParagraph(P paragraph, Map<String, Object> context) {
        List<Text> texts = new ArrayList<>();

        new TraversalUtil(paragraph, new Callback() {
            @Override
            public List<Object> apply(Object o) {
                Object u = XmlUtils.unwrap(o);
                if (u instanceof Text t) {
                    texts.add(t);
                }
                return null;
            }

            @Override
            public boolean shouldTraverse(Object o) {
                return true;
            }

            @Override
            public void walkJAXBElements(Object parent) {
                List<Object> children = getChildren(parent);
                if (children != null) {
                    for (Object o : children) {
                        Object u = XmlUtils.unwrap(o);
                        apply(u);
                        walkJAXBElements(u);
                    }
                }
            }

            @Override
            public List<Object> getChildren(Object o) {
                return TraversalUtil.getChildrenImpl(o);
            }
        });

        if (texts.isEmpty()) return;

        StringBuilder original = new StringBuilder();
        for (Text t : texts) {
            original.append(t.getValue());
        }
        String originalText = original.toString();

        if (!originalText.contains("{{")) return;

        String renderedText = templateEngine.render(originalText, context);

        if (renderedText.equals(originalText)) return;

        boolean first = true;
        for (Text t : texts) {
            if (first) {
                t.setValue(renderedText);
                first = false;
            } else {
                t.setValue("");
            }
        }
    }

    @SuppressWarnings("unchecked")
    private Map<String, Object> castToMap(Object item) {
        if (item instanceof Map<?, ?> m) {
            return (Map<String, Object>) m;
        }
        // primitive -> dùng {{.}} nếu cần
        return Map.of(".", item);
    }
}
