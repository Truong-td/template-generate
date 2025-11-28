package com.truongtd.templategenerate.service.impl;

import com.truongtd.templategenerate.dto.TemplateDataDto;
import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.helper.TemplateContextBuilder;
import com.truongtd.templategenerate.mapper.TemplateMapper;
import com.truongtd.templategenerate.request.CreateTemplateRequest;
import com.truongtd.templategenerate.service.TemplateService;
import jakarta.annotation.PostConstruct;
import lombok.val;
import org.apache.commons.compress.utils.IOUtils;
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
import java.io.FileInputStream;
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
    private final TableDataService tableDataService;

    //    // Block list/object: {{students}} ... {{/students}}
//    private static final Pattern BLOCK_START_PATTERN =
//            Pattern.compile("^\\s*\\{\\{(\\w+)}}\\s*$");
//
//    // Conditional block: {{?user}} ... {{/user}}
//    private static final Pattern COND_START_PATTERN =
//            Pattern.compile("^\\s*\\{\\{\\?(\\w+)}}\\s*$");
//
//    // Block end: {{/students}}
//    private static final Pattern BLOCK_END_PATTERN =
//            Pattern.compile("^\\s*\\{\\{/(\\w+)}}\\s*$");
//
//    // Row template trong bảng: {{students.name}}
//    private static final Pattern ROW_LIST_PLACEHOLDER =
//            Pattern.compile("\\{\\{(\\w+)\\.[^}]+}}");
//
//    // Factory tạo node wml
//    private static final ObjectFactory WML_FACTORY = new ObjectFactory();
    public static final Pattern BLOCK_START = Pattern.compile("\\{\\{\\?(.*?)}}");
    public static final Pattern BLOCK_END   = Pattern.compile("\\{\\{/(.*?)}}");
    public static final Pattern SCALAR      = Pattern.compile("\\{\\{([^{}]+)}}");

    public TemplateServiceImpl(TableDataService tableDataService) {
        this.tableDataService = tableDataService;
    }

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
    public byte[] generateDocx(GenerateTemplateRequest request) throws Exception {

        try {
            TemplateDataDto templateDataDto = JsonUtils.parse(request);

            // 1. Load file mẫu generateTemplate.docx từ resources/templates
            WordprocessingMLPackage wordMLPackage = loadTemplate();

            // 2. Build root context cho textData + tableData
            Map<String, Object> context = contextBuilder.buildRootContext(templateDataDto);

            // 3. Xử lý FlexData: {{historyFormation}}
            processFlexData(wordMLPackage, templateDataDto.getFlexData());

            // 5. Xử lý TableData: lặp danh sách students, subjects...
            tableDataService.processTableData(wordMLPackage, context);

            // 4. Xử lý TextData: scalar + block ẩn/hiện ({{?key}}...{{/key}})
            processTextBlocks(wordMLPackage, context);

            cleanupEmptyParagraphs(wordMLPackage);

            // 6. Xuất docx ra mảng byte
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wordMLPackage.save(out);
            return out.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Error generating template", e);
        }

    }

    @SuppressWarnings("unchecked")
    private void processFlexData(WordprocessingMLPackage pkg, Map<String, Object> flexData) throws Exception {
        if (flexData == null) return;

        for (Map.Entry<String, Object> e : flexData.entrySet()) {
            String key = e.getKey();      // ví dụ: historyFormation, collateralInfo, timeline...
            Object val = e.getValue();

            if (val instanceof List<?>) {
                processOneFlexBlock(pkg, key, (List<Object>) val);
            }
        }
    }

    private void processOneFlexBlock(WordprocessingMLPackage pkg,
                                     String placeholderKey,
                                     List<Object> blocks) throws Exception {
        MainDocumentPart main = pkg.getMainDocumentPart();
        Body body = main.getContents().getBody();

        String tag = "{{" + placeholderKey + "}}";

        // tìm paragraph chứa {{placeholderKey}}
        P placeholderPara = null;
        for (Object obj : body.getContent()) {
            Object u = XmlUtils.unwrap(obj);
            if (u instanceof P) {
                P p = (P) u;
                if (getParagraphText(p).contains(tag)) {
                    placeholderPara = p;
                    break;
                }
            }
        }
        if (placeholderPara == null) return;

        int insertIndex = body.getContent().indexOf(placeholderPara);
        body.getContent().remove(placeholderPara);

        for (Object o : blocks) {
            if (!(o instanceof Map)) continue;
            Map<String, Object> block = (Map<String, Object>) o;
            String type = (String) block.get("type");

            switch (type) {
                case "text":
                    body.getContent().add(insertIndex++, createTextParagraph(block));
                    break;
                case "table":
                    body.getContent().add(insertIndex++, createTableFromFlex(block));
                    break;
                case "image":
                    body.getContent().add(insertIndex++, createImageParagraph(pkg, block));
                    break;
                default:
                    break;
            }
        }
    }

    @SuppressWarnings("unchecked")
    private P createTextParagraph(Map<String, Object> block) {
        Map<String, Object> textData = (Map<String, Object>) block.get("textData");
        String data = textData != null ? (String) textData.get("data") : "";
        Boolean isBold = textData != null && Boolean.TRUE.equals(textData.get("isBold"));

        P p = new P();
        R r = new R();
        Text t = new Text();
        t.setValue(data);
        r.getContent().add(t);

        if (isBold) {
            RPr rPr = new RPr();
            BooleanDefaultTrue b = new BooleanDefaultTrue();
            rPr.setB(b);
            r.setRPr(rPr);
        }

        p.getContent().add(r);
        return p;
    }

    @SuppressWarnings("unchecked")
    private Tbl createTableFromFlex(Map<String, Object> block) {
        Map<String, Object> tableData = (Map<String, Object>) block.get("tableData");
        List<List<Map<String, Object>>> matrix =
                (List<List<Map<String, Object>>>) tableData.get("data");
        List<Double> colSizes = (List<Double>) tableData.get("colSizes");

        Tbl tbl = new Tbl();

        for (List<Map<String, Object>> rowData : matrix) {
            Tr tr = new Tr();
            for (int i = 0; i < rowData.size(); i++) {
                Map<String, Object> cellData = rowData.get(i);
                String text = (String) cellData.get("data");
                Boolean isBold = Boolean.TRUE.equals(cellData.get("isBold"));

                Tc tc = new Tc();
                P p = new P();
                R r = new R();
                Text t = new Text();
                t.setValue(text != null ? text : "");
                r.getContent().add(t);

                if (isBold) {
                    RPr rPr = new RPr();
                    rPr.setB(new BooleanDefaultTrue());
                    r.setRPr(rPr);
                }

                p.getContent().add(r);
                tc.getContent().add(p);
                tr.getContent().add(tc);
            }
            tbl.getContent().add(tr);
        }

        // nếu cần set width theo colSizes anh có thể bổ sung TcPr/TblGrid ở đây
        return tbl;
    }

    @SuppressWarnings("unchecked")
    private P createImageParagraph(WordprocessingMLPackage pkg, Map<String, Object> block) throws Exception {
        Map<String, Object> imageData = (Map<String, Object>) block.get("imageData");
        if (imageData == null) return new P();

        String bucket = (String) imageData.get("bucket");
        String path = (String) imageData.get("path");
        // TODO: anh tự implement load bytes từ S3 hoặc file local
        byte[] bytes = loadImageBytes(bucket, path);

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage
                .createImagePart(pkg, bytes);

        Inline inline = imagePart.createImageInline("history-img", "history-img",
                0, 1, 6000, false);

        R r = new R();
        Drawing drawing = new Drawing();
        drawing.getAnchorOrInline().add(inline);
        r.getContent().add(drawing);
        P p = new P();
        p.getContent().add(r);
        return p;
    }

    private byte[] loadImageBytes(String bucket, String path) {
        // ví dụ tạm: đọc từ local cho dễ debug
        try (InputStream is = new FileInputStream("/tmp/" + path.substring(path.lastIndexOf('/') + 1))) {
            return IOUtils.toByteArray(is);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void processTextBlocks(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {
        List<Object> paragraphs = getAllParagraphs(pkg);

        // Xử lý block theo dạng: start -> các paragraph bên trong -> end
        boolean insideBlock = false;
        String currentKey = null;
        List<P> buffer = new ArrayList<>();

        for (Iterator<Object> it = paragraphs.iterator(); it.hasNext(); ) {
            Object obj = it.next();
            P p = (P) XmlUtils.unwrap(obj);

            String text = getParagraphText(p);
            Matcher startM = BLOCK_START.matcher(text);
            Matcher endM   = BLOCK_END.matcher(text);

            if (!insideBlock && startM.find()) {
                // bắt đầu block
                insideBlock = true;
                currentKey = startM.group(1).trim();
                buffer.clear();
                buffer.add(p); // paragraph chứa {{?key}}
                continue;
            }

            if (insideBlock) {
                buffer.add(p);
                if (endM.find()) {
                    // kết thúc block, quyết định show/hide
                    boolean show = isTruthy(resolveKey(root, currentKey));
                    if (show) {
                        // render từng paragraph trong buffer: remove {{?key}}, {{/key}} và scalar
                        for (P blockP : buffer) {
                            String t = getParagraphText(blockP);
                            t = t.replace("{{?" + currentKey + "}}", "")
                                    .replace("{{/" + currentKey + "}}", "");

                            t = replaceScalars(t, root, currentKey);

                            // Nếu sau khi xoá marker + replace mà paragraph trống hoàn toàn -> xoá hẳn
                            if (t == null || t.trim().isEmpty()) {
                                deleteParagraph(pkg, blockP);
                            } else {
                                setParagraphText(blockP, t);
                            }
                        }
                    } else {
                        // xóa toàn bộ block
                        for (P blockP : buffer) {
                            deleteParagraph(pkg, blockP);
                        }
                    }

                    insideBlock = false;
                    currentKey = null;
                    buffer.clear();
                }
                continue;
            }

            // ngoài block: chỉ thay scalar bình thường
            String replaced = replaceScalars(text, root, null);
            setParagraphText(p, replaced);
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
                prevEmpty = false; // reset nếu gặp bảng / hình / gì khác
                continue;
            }

            P p = (P) o;
            String txt = getParagraphText(p).trim();
            boolean isEmpty = txt.isEmpty();

            if (isEmpty) {
                if (prevEmpty) {
                    // 2 paragraph trống liên tiếp -> bỏ cái thứ 2 trở đi
                    it.remove();
                } else {
                    prevEmpty = true;
                }
            } else {
                prevEmpty = false;
            }
        }
    }
    private List<Object> getAllParagraphs(WordprocessingMLPackage pkg) {
        try {
            return pkg.getMainDocumentPart()
                    .getJAXBNodesViaXPath("//w:p", true);
        } catch (Exception e) {
            throw new RuntimeException("Cannot read paragraphs", e);
        }
    }

    private void setParagraphText(P p, String newText) {
        // clear all run and replace bằng 1 run mới
        p.getContent().clear();
        R run = new R();
        Text text = new Text();
        text.setValue(newText);
        run.getContent().add(text);
        p.getContent().add(run);
    }

    private void deleteParagraph(WordprocessingMLPackage pkg, P p) throws Docx4JException {
        Body body = pkg.getMainDocumentPart().getContents().getBody();
        body.getContent().remove(p);
    }

    // isTruthy: cho Trường hợp 1 & 2
    private boolean isTruthy(Object value) {
        if (value == null) return false;
        if (value instanceof Boolean) return (Boolean) value;
        if (value instanceof String) return !((String) value).trim().isEmpty();
        if (value instanceof Collection) return !((Collection<?>) value).isEmpty();
        if (value instanceof Map) return !((Map<?, ?>) value).isEmpty();
        return true;
    }

    private Object resolveKey(Map<String, Object> root, String key) {
        // hỗ trợ nested: user.name, application.name...
        String[] parts = key.split("\\.");
        Object current = root;
        for (String part : parts) {
            if (!(current instanceof Map)) return null;
            current = ((Map<?, ?>) current).get(part);
            if (current == null) return null;
        }
        return current;
    }

    // Lấy paragraph text (ghép các run)
    private String getParagraphText(P p) {
        StringBuilder sb = new StringBuilder();
        for (Object o : p.getContent()) {
            Object u = XmlUtils.unwrap(o);
            if (u instanceof R) {
                for (Object c : ((R) u).getContent()) {
                    Object cu = XmlUtils.unwrap(c);
                    if (cu instanceof Text) {
                        sb.append(((Text) cu).getValue());
                    }
                }
            }
        }
        return sb.toString();
    }

    private String replaceScalars(String text, Map<String, Object> root, String blockKey) {
        // nếu nằm trong block {{?user}} thì context gốc cho {{name}} là object user
        Map<String, Object> ctx = root;
        if (blockKey != null) {
            Object obj = resolveKey(root, blockKey);
            if (obj instanceof Map) {
                //noinspection unchecked
                ctx = (Map<String, Object>) obj;
            }
        }

        Matcher m = SCALAR.matcher(text);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String key = m.group(1).trim();
            Object value = resolveKey(ctx, key);
            m.appendReplacement(sb, Matcher.quoteReplacement(value != null ? String.valueOf(value) : ""));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    private WordprocessingMLPackage loadTemplate() throws Docx4JException, IOException {
        try (InputStream is = getClass().getResourceAsStream("/templates/template-report.docx")) {
            if (is == null) {
                throw new IllegalStateException("Không tìm thấy template-report.docx trong /templates");
            }
            return WordprocessingMLPackage.load(is);
        }
    }
}
