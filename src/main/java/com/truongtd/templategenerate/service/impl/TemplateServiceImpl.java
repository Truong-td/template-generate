package com.truongtd.templategenerate.service.impl;

import com.truongtd.templategenerate.dto.TemplateDataDto;
import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.helper.TemplateContextBuilder;
import com.truongtd.templategenerate.mapper.TemplateMapper;
import com.truongtd.templategenerate.request.CreateTemplateRequest;
import com.truongtd.templategenerate.service.TemplateService;
import jakarta.annotation.PostConstruct;
import lombok.val;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.docx4j.TraversalUtil.Callback;

@Service
public class TemplateServiceImpl implements TemplateService {

    private static final Logger log = LoggerFactory.getLogger(TemplateServiceImpl.class);

    private final DocxTemplateEngine templateEngine = new DocxTemplateEngine();
    private final TemplateContextBuilder contextBuilder = new TemplateContextBuilder();

    // Block list/object: {{students}} ... {{/students}}
    private static final Pattern BLOCK_START_PATTERN =
            Pattern.compile("^\\s*\\{\\{(\\w+)}}\\s*$");

    // Conditional block: {{?user}} ... {{/user}}
    private static final Pattern COND_START_PATTERN =
            Pattern.compile("^\\s*\\{\\{\\?(\\w+)}}\\s*$");

    // Block end: {{/students}}
    private static final Pattern BLOCK_END_PATTERN =
            Pattern.compile("^\\s*\\{\\{/(\\w+)}}\\s*$");

    // Row template trong bảng: {{students.name}}
    private static final Pattern ROW_LIST_PLACEHOLDER =
            Pattern.compile("\\{\\{(\\w+)\\.[^}]+}}");

    // Factory tạo node wml
    private static final ObjectFactory WML_FACTORY = new ObjectFactory();

    // FlexData markers
    private static final Pattern IMG_PATTERN =
            Pattern.compile("\\{\\{IMG:([^}]+)}}");
    private static final Pattern HTML_PATTERN =
            Pattern.compile("\\{\\{HTML:([^}]+)}}");
    private static final Pattern TABLE2D_PATTERN =
            Pattern.compile("\\{\\{TABLE:([^}]+)}}");

    // HTTP client để tải ảnh
    private static final HttpClient HTTP_CLIENT = HttpClient.newHttpClient();

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

        // 1. TableData -> 1 bảng nhiều dòng
        processTablesForList(mainDocumentPart, rootContext);

        // 2. Block list/object + conditional + scalar cho toàn document
        processDocument(mainDocumentPart, rootContext);

        // 3. FlexData: IMG / HTML / TABLE2D
        processFlex(mainDocumentPart, rootContext, wordMLPackage);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wordMLPackage.save(out);
        return out.toByteArray();
    }

    private void processFlex(MainDocumentPart mainDocumentPart,
                             Map<String, Object> rootContext,
                             WordprocessingMLPackage wordMLPackage) throws Exception {

        Document doc = (Document) mainDocumentPart.getJaxbElement();
        List<Object> bodyContent = doc.getBody().getContent();

        for (int i = 0; i < bodyContent.size(); i++) {
            Object el = bodyContent.get(i);
            Object u = XmlUtils.unwrap(el);

            if (!(u instanceof P p)) {
                continue;
            }

            String text = getParagraphText(p);
            if (text == null) continue;

            // IMG
            Matcher imgM = IMG_PATTERN.matcher(text);
            if (imgM.find()) {
                String key = imgM.group(1).trim(); // ví dụ "note.image"
                Object val = resolveKey(key, rootContext);
                if (val != null) {
                    try {
                        P imgPara = createImageParagraph(wordMLPackage, val.toString());
                        bodyContent.set(i, imgPara);
                    } catch (Exception e) {
                        log.error("Lỗi chèn ảnh cho key {}", key, e);
                    }
                } else {
                    log.warn("Không tìm thấy dữ liệu ảnh cho key {}", key);
                }
                continue;
            }

            // HTML
            Matcher htmlM = HTML_PATTERN.matcher(text);
            if (htmlM.find()) {
                String key = htmlM.group(1).trim(); // ví dụ "note.html"
                Object val = resolveKey(key, rootContext);
                if (val != null) {
                    try {
                        List<Object> htmlNodes = convertHtmlToNodes(mainDocumentPart, val.toString());
                        // Xoá paragraph chứa marker, chèn HTML tại vị trí đó
                        bodyContent.remove(i);
                        bodyContent.addAll(i, htmlNodes);
                        i += htmlNodes.size() - 1;
                    } catch (Exception e) {
                        log.error("Lỗi chèn HTML cho key {}", key, e);
                    }
                } else {
                    log.warn("Không tìm thấy dữ liệu HTML cho key {}", key);
                }
                continue;
            }

            // TABLE2D
            Matcher tblM = TABLE2D_PATTERN.matcher(text);
            if (tblM.find()) {
                String key = tblM.group(1).trim(); // ví dụ "note.table"
                Object val = resolveKey(key, rootContext);
                if (val instanceof List<?>) {
                    try {
                        Tbl tbl = createTableFrom2D((List<?>) val);
                        bodyContent.set(i, tbl);
                    } catch (Exception e) {
                        log.error("Lỗi chèn TABLE cho key {}", key, e);
                    }
                } else {
                    log.warn("Dữ liệu TABLE cho key {} không phải List", key);
                }
            }
        }
    }

    private P createImageParagraph(WordprocessingMLPackage wordMLPackage, String url) throws Exception {
        byte[] bytes = downloadImage(url);
        if (bytes == null || bytes.length == 0) {
            throw new IllegalStateException("Không tải được ảnh từ URL: " + url);
        }

        BinaryPartAbstractImage imagePart =
                BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        // Kích thước ảnh (đơn vị EMU) – tạm approx 400x300 px
        long cx = 400L * 9525L;
        long cy = 300L * 9525L;

        Inline inline = imagePart.createImageInline(
                "flex-img", "Flex image", 0, 1, cx, cy, false);

        Drawing drawing = WML_FACTORY.createDrawing();
        drawing.getAnchorOrInline().add(inline);

        R run = WML_FACTORY.createR();
        run.getContent().add(drawing);

        P p = WML_FACTORY.createP();
        p.getContent().add(run);
        return p;
    }

    private byte[] downloadImage(String url) throws Exception {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .GET()
                .build();

        HttpResponse<byte[]> response =
                HTTP_CLIENT.send(request, HttpResponse.BodyHandlers.ofByteArray());

        if (response.statusCode() >= 200 && response.statusCode() < 300) {
            return response.body();
        } else {
            log.warn("Tải ảnh thất bại {} - status {}", url, response.statusCode());
            return null;
        }
    }

    private List<Object> convertHtmlToNodes(MainDocumentPart mainDocumentPart, String html) throws Exception {
        WordprocessingMLPackage wordMLPackage = loadTemplate();
        XHTMLImporterImpl importer =
                new XHTMLImporterImpl(wordMLPackage);
        // html có thể là đoạn <p>... hoặc <div>...
        return importer.convert(html, null);
    }

    @SuppressWarnings("unchecked")
    private Tbl createTableFrom2D(List<?> rows) {
        Tbl tbl = WML_FACTORY.createTbl();

        for (Object rowObj : rows) {
            if (!(rowObj instanceof List<?> cols)) {
                continue;
            }
            Tr tr = WML_FACTORY.createTr();

            for (Object col : cols) {
                Tc tc = WML_FACTORY.createTc();
                P p = WML_FACTORY.createP();
                Text t = WML_FACTORY.createText();
                t.setValue(col != null ? col.toString() : "");
                R r = WML_FACTORY.createR();
                r.getContent().add(t);
                p.getContent().add(r);
                tc.getContent().add(p);
                tr.getContent().add(tc);
            }

            tbl.getContent().add(tr);
        }

        return tbl;
    }

    private WordprocessingMLPackage loadTemplate() throws Docx4JException, IOException {
        try (InputStream is = getClass().getResourceAsStream("/templates/template-report.docx")) {
            if (is == null) {
                throw new IllegalStateException("Không tìm thấy template-report.docx trong /templates");
            }
            return WordprocessingMLPackage.load(is);
        }
    }

    /* =========================================================
       1) TABLE DATA – 1 bảng nhiều dòng
       ========================================================= */

    @SuppressWarnings("unchecked")
    private void processTablesForList(MainDocumentPart mainDocumentPart,
                                      Map<String, Object> rootContext) {

        Document wmlDoc = (Document) mainDocumentPart.getJaxbElement();
        List<Object> bodyContent = wmlDoc.getBody().getContent();

        for (Object el : bodyContent) {
            Object u = XmlUtils.unwrap(el);
            if (!(u instanceof Tbl tbl)) {
                continue;
            }

            List<Tr> rows = new ArrayList<>();
            for (Object rowObj : tbl.getContent()) {
                Object ru = XmlUtils.unwrap(rowObj);
                if (ru instanceof Tr tr) {
                    rows.add(tr);
                }
            }

            if (rows.isEmpty()) continue;

            int templateIndex = -1;
            String listName = null;

            // tìm row template
            for (int i = 0; i < rows.size(); i++) {
                Tr tr = rows.get(i);
                String rowText = getRowText(tr);
                if (rowText == null) continue;

                Matcher m = ROW_LIST_PLACEHOLDER.matcher(rowText);
                if (m.find()) {
                    templateIndex = i;
                    listName = m.group(1); // ví dụ "students"
                    break;
                }
            }

            if (templateIndex == -1 || listName == null) {
                continue; // bảng không có row template
            }

            log.debug("Table: template row index={} cho list={}", templateIndex, listName);

            Object data = rootContext.get(listName);
            if (!(data instanceof List<?> dataList) || dataList.isEmpty()) {
                // không có data -> xoá row template, giữ header
                tbl.getContent().remove(templateIndex);
                log.debug("List {} không có data, xoá row template", listName);
                continue;
            }

            Tr templateRow = rows.get(templateIndex);
            // xoá row template
            tbl.getContent().remove(templateIndex);

            int insertPos = templateIndex;

            for (Object item : dataList) {
                Map<String, Object> itemMap = castToMap(item);

                // context = root + listName = item
                Map<String, Object> combinedContext = new HashMap<>(rootContext);
                combinedContext.put(listName, itemMap);

                Tr newRow = (Tr) XmlUtils.deepCopy(templateRow);
                replaceScalarsInRow(newRow, combinedContext);
                tbl.getContent().add(insertPos++, newRow);
            }

            log.debug("Nhân bản row template cho list {} với {} dòng", listName, ((List<?>) data).size());
        }
    }

    private String getRowText(Tr row) {
        List<Text> texts = new ArrayList<>();

        new TraversalUtil(row, new Callback() {
            @Override
            public List<Object> apply(Object o) {
                Object u = XmlUtils.unwrap(o);
                if (u instanceof Text t) {
                    texts.add(t);
                }
                return null;
            }

            @Override
            public boolean shouldTraverse(Object o) { return true; }

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

    private void replaceScalarsInRow(Tr row, Map<String, Object> context) {
        List<Text> texts = new ArrayList<>();

        new TraversalUtil(row, new Callback() {
            @Override
            public List<Object> apply(Object o) {
                Object u = XmlUtils.unwrap(o);
                if (u instanceof Text t) {
                    texts.add(t);
                }
                return null;
            }

            @Override
            public boolean shouldTraverse(Object o) { return true; }

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

    /* =========================================================
       2) BLOCK & CONDITIONAL – toàn document
       ========================================================= */

    @SuppressWarnings("unchecked")
    private void processDocument(MainDocumentPart mainDocumentPart,
                                 Map<String, Object> rootContext) {

        Document wmlDoc = (Document) mainDocumentPart.getJaxbElement();
        List<Object> originalContent = wmlDoc.getBody().getContent();
        List<Object> newContent = new ArrayList<>();

        for (int i = 0; i < originalContent.size(); i++) {
            Object el = originalContent.get(i);
            Object unwrapped = XmlUtils.unwrap(el);

            if (!(unwrapped instanceof P p)) {
                newContent.add(XmlUtils.deepCopy(el));
                continue;
            }

            String paraText = getParagraphText(p);
            if (paraText == null) {
                P cloned = (P) XmlUtils.deepCopy(p);
                newContent.add(cloned);
                continue;
            }

            String trimmed = paraText.trim();

            /* ---------- 2.1 CONDITIONAL: {{?user}} ... {{/user}} ---------- */
            Matcher condStart = COND_START_PATTERN.matcher(trimmed);
            if (condStart.matches()) {
                String blockName = condStart.group(1);
                int endIndex = findBlockEnd(originalContent, blockName, i + 1);
                if (endIndex == -1) {
                    // treat as normal paragraph
                    P cloned = (P) XmlUtils.deepCopy(p);
                    replaceScalarsInParagraph(cloned, rootContext);
                    newContent.add(cloned);
                    continue;
                }

                Object value = rootContext.get(blockName);
                boolean shouldRender = evaluateCondition(value);

                if (shouldRender) {
                    List<Object> body = originalContent.subList(i + 1, endIndex);
                    for (Object bodyEl : body) {
                        Object copy = XmlUtils.deepCopy(bodyEl);
                        Object bodyUnwrapped = XmlUtils.unwrap(copy);
                        if (bodyUnwrapped instanceof P bodyPara) {
                            replaceScalarsInParagraph(bodyPara, rootContext);
                        }
                        newContent.add(copy);
                    }
                }
                i = endIndex; // skip to end
                continue;
            }

            /* ---------- 2.2 BLOCK LIST/OBJECT: {{students}} ... {{/students}} ---------- */
            Matcher startMatcher = BLOCK_START_PATTERN.matcher(trimmed);
            if (startMatcher.matches()) {
                String blockName = startMatcher.group(1);

                int endIndex = findBlockEnd(originalContent, blockName, i + 1);
                if (endIndex == -1) {
                    P cloned = (P) XmlUtils.deepCopy(p);
                    replaceScalarsInParagraph(cloned, rootContext);
                    newContent.add(cloned);
                    continue;
                }

                List<Object> body = originalContent.subList(i + 1, endIndex);
                Object blockData = rootContext.get(blockName);

                if (blockData instanceof List<?> list) {
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
                    Map<String, Object> subContext = (Map<String, Object>) map;
                    for (Object bodyEl : body) {
                        Object copy = XmlUtils.deepCopy(bodyEl);
                        Object bodyUnwrapped = XmlUtils.unwrap(copy);
                        if (bodyUnwrapped instanceof P bodyPara) {
                            replaceScalarsInParagraph(bodyPara, subContext);
                        }
                        newContent.add(copy);
                    }
                }
                // nếu không có data phù hợp thì bỏ luôn block
                i = endIndex;
                continue;
            }

            /* ---------- 2.3 END mồ côi -> bỏ ---------- */
            Matcher endMatcher = BLOCK_END_PATTERN.matcher(trimmed);
            if (endMatcher.matches()) {
                continue;
            }

            /* ---------- 2.4 Paragraph thường ---------- */
            P cloned = (P) XmlUtils.deepCopy(p);
            replaceScalarsInParagraph(cloned, rootContext);
            newContent.add(cloned);
        }

        originalContent.clear();
        originalContent.addAll(newContent);
    }

    private int findBlockEnd(List<Object> content, String blockName, int fromIndex) {
        for (int j = fromIndex; j < content.size(); j++) {
            Object elEnd = content.get(j);
            Object uEnd = XmlUtils.unwrap(elEnd);
            if (uEnd instanceof P pEnd) {
                String endText = getParagraphText(pEnd);
                if (endText != null) {
                    Matcher endMatcher = BLOCK_END_PATTERN.matcher(endText.trim());
                    if (endMatcher.matches() && blockName.equals(endMatcher.group(1))) {
                        return j;
                    }
                }
            }
        }
        return -1;
    }

    private boolean evaluateCondition(Object value) {
        if (value == null) return false;
        if (value instanceof Boolean b) return b;
        if (value instanceof Collection<?> c) return !c.isEmpty();
        if (value instanceof Map<?, ?> m) return !m.isEmpty();
        // primitive hoặc object bất kỳ khác null => true
        return true;
    }

    /* =========================================================
       Helpers: paragraph scalar
       ========================================================= */

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
            public boolean shouldTraverse(Object o) { return true; }

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
        for (Text t : texts) sb.append(t.getValue());
        return sb.toString();
    }

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
            public boolean shouldTraverse(Object o) { return true; }

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
        for (Text t : texts) original.append(t.getValue());
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
        return Map.of(".", item);
    }

    @SuppressWarnings("unchecked")
    private Object resolveKey(String key, Map<String, Object> context) {
        if (!key.contains(".")) {
            return context.get(key);
        }
        String[] parts = key.split("\\.");
        Object current = context;
        for (String part : parts) {
            if (!(current instanceof Map)) return null;
            current = ((Map<String, Object>) current).get(part);
            if (current == null) return null;
        }
        return current;
    }
}
