package com.truongtd.templategenerate.service.impl;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.truongtd.templategenerate.dto.TemplateDataDto;
import com.truongtd.templategenerate.request.GenerateTemplateRequest;
import com.truongtd.templategenerate.service.TemplateService;
import com.truongtd.templategenerate.util.JsonUtils;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.wml.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class TemplateServiceImpl implements TemplateService {

    @Value("${template.image.base-dir:}")   // optional base dir
    private String imageBaseDir;

    private static final Pattern BLOCK_START =
            Pattern.compile("\\{\\{\\?(.*?)}}");
    private static final Pattern BLOCK_END =
            Pattern.compile("\\{\\{/(.*?)}}");
    private static final Pattern SCALAR =
            Pattern.compile("\\{\\{([^{}]+)}}");
    private static final Pattern LIST_IN_ROW =
            Pattern.compile("\\{\\{([a-zA-Z0-9_]+)\\.[^}]+}}");

    @Override
    public byte[] generateDocx(GenerateTemplateRequest request) {
        try {
            TemplateDataDto data = JsonUtils.parse(request);

            WordprocessingMLPackage pkg = WordprocessingMLPackage.load(
                    getClass().getResourceAsStream("/templates/template-report.docx"));

            Map<String, Object> context = buildRootContext(data);

            // 1. FlexData: thay {{key}} bằng text / table / image
            processFlexData(pkg, data.getFlexData());

            // 2. TableData: lặp các bảng list + custom "Danh sách môn học"
            processTableData(pkg, context);
            fillSubjectsText(pkg.getMainDocumentPart()
                    .getContents().getBody(), context);

            // 3. TextData: scalar + block {{?key}}...{{/key}}
            processTextBlocks(pkg, context);

            // 4. Dọn các paragraph rỗng dư thừa
            cleanupEmptyParagraphs(pkg);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            pkg.save(out);
            return out.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Error generating template", e);
        }
    }

    // ===================== ROOT CONTEXT =====================

    private Map<String, Object> buildRootContext(TemplateDataDto data) {
        Map<String, Object> root = new HashMap<>();
        if (data.getTextData() != null) {
            root.putAll(data.getTextData());
        }
        if (data.getTableData() != null) {
            root.putAll(data.getTableData());
        }
        // KHÔNG put flexData vào root, để flex xử lý riêng
        return root;
    }

    // ===================== FLEX DATA (generic nhiều key) =====================

    @SuppressWarnings("unchecked")
    private void processFlexData(WordprocessingMLPackage pkg,
                                 Map<String, Object> flexData) throws Exception {
        if (flexData == null || flexData.isEmpty()) return;

        for (Map.Entry<String, Object> e : flexData.entrySet()) {
            String key = e.getKey();      // ví dụ: historyFormation, note, ...
            Object val = e.getValue();

            List<Object> blocks = null;

            // Case 1: đã là mảng block (kiểu historyFormation cũ)
            if (val instanceof List) {
                blocks = (List<Object>) val;
            }
            // Case 2: object đơn giản { text, table, image }
            else if (val instanceof Map) {
                blocks = convertSimpleFlexObjectToBlocks((Map<String, Object>) val);
            }
            // Case 3: chỉ là 1 chuỗi text
            else if (val instanceof String) {
                blocks = new ArrayList<>();
                Map<String, Object> block = new HashMap<>();
                block.put("type", "text");
                Map<String, Object> textData = new HashMap<>();
                textData.put("data", val);
                textData.put("isBold", false);
                block.put("textData", textData);
                blocks.add(block);
            }

            if (blocks == null || blocks.isEmpty()) {
                continue;
            }

            processOneFlexBlock(pkg, key, blocks); // dùng lại hàm cũ
        }
    }
    @SuppressWarnings("unchecked")
    private List<Object> convertSimpleFlexObjectToBlocks(Map<String, Object> obj) {
        List<Object> blocks = new ArrayList<>();

        // --- text ---
        Object textVal = obj.get("text");
        if (textVal != null) {
            Map<String, Object> block = new HashMap<>();
            block.put("type", "text");

            Map<String, Object> textData = new HashMap<>();
            textData.put("data", String.valueOf(textVal));
            textData.put("isBold", false); // tuỳ anh

            block.put("textData", textData);
            blocks.add(block);
        }

        // --- table: [[...], [...]] ---
        Object tableVal = obj.get("table");
        if (tableVal instanceof List) {
            List<List<?>> rawTable = (List<List<?>>) tableVal;

            List<List<Map<String, Object>>> matrix = new ArrayList<>();
            for (int rowIdx = 0; rowIdx < rawTable.size(); rowIdx++) {
                List<?> row = rawTable.get(rowIdx);
                List<Map<String, Object>> rowOut = new ArrayList<>();
                for (Object cellVal : row) {
                    Map<String, Object> cell = new HashMap<>();
                    cell.put("data", cellVal != null ? String.valueOf(cellVal) : "");
                    // ví dụ: dòng đầu là header => in đậm
                    cell.put("isBold", rowIdx == 0);
                    rowOut.add(cell);
                }
                matrix.add(rowOut);
            }

            Map<String, Object> block = new HashMap<>();
            block.put("type", "table");

            Map<String, Object> tableData = new HashMap<>();
            tableData.put("data", matrix);
            block.put("tableData", tableData);

            blocks.add(block);
        }

        // --- image: URL string ---
        Object imageVal = obj.get("image");
        if (imageVal != null) {
            Map<String, Object> block = new HashMap<>();
            block.put("type", "image");

            Map<String, Object> imageData = new HashMap<>();
            imageData.put("bucket", null);              // không dùng bucket
            imageData.put("path", String.valueOf(imageVal)); // URL hoặc file path
            block.put("imageData", imageData);

            blocks.add(block);
        }

        return blocks;
    }

    private void processOneFlexBlock(WordprocessingMLPackage pkg,
                                     String placeholderKey,
                                     List<Object> blocks) throws Exception {
        MainDocumentPart main = pkg.getMainDocumentPart();
        Body body = main.getContents().getBody();

        String tag = "{{" + placeholderKey + "}}";

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
            @SuppressWarnings("unchecked")
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
            rPr.setB(new BooleanDefaultTrue());
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
        // List<Double> colSizes = (List<Double>) tableData.get("colSizes"); // dùng nếu cần

        Tbl tbl = new Tbl();

        for (List<Map<String, Object>> rowData : matrix) {
            Tr tr = new Tr();
            for (Map<String, Object> cellData : rowData) {
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
        return tbl;
    }

    @SuppressWarnings("unchecked")
    private P createImageParagraph(WordprocessingMLPackage pkg,
                                   Map<String, Object> block) throws Exception {
        Map<String, Object> imageData = (Map<String, Object>) block.get("imageData");
        if (imageData == null) return new P();

        String bucket = (String) imageData.get("bucket");
        String path = (String) imageData.get("path");

        byte[] bytes = loadImageBytes(bucket, path);
        if (bytes == null || bytes.length == 0) return new P();

        BinaryPartAbstractImage imagePart =
                BinaryPartAbstractImage.createImagePart(pkg, bytes);

        Inline inline = imagePart.createImageInline(
                "flex-img", "flex-img", 0, 1, 6000, false);

        Drawing drawing = new Drawing();
        drawing.getAnchorOrInline().add(inline);

        R r = new R();
        r.getContent().add(drawing);

        P p = new P();
        p.getContent().add(r);
        return p;
    }

    private byte[] loadImageBytes(String bucket, String path) {
        if (path == null || path.trim().isEmpty()) {
            return new byte[0];
        }
        path = path.trim();

        try {
            // 1) URL
            if (path.startsWith("http://") || path.startsWith("https://")) {
                try (InputStream is = new URL(path).openStream()) {
                    return IOUtils.toByteArray(is);
                }
            }

            // 2) classpath:...
            if (path.startsWith("classpath:")) {
                String cp = path.substring("classpath:".length());
                String resourcePath = cp.startsWith("/") ? cp : "/" + cp;
                try (InputStream is = getClass().getResourceAsStream(resourcePath)) {
                    if (is == null) {
                        throw new FileNotFoundException("Classpath resource not found: " + cp);
                    }
                    return IOUtils.toByteArray(is);
                }
            }

            // 3) file hệ thống
            Path filePath;
            if (Paths.get(path).isAbsolute()) {
                filePath = Paths.get(path);
            } else {
                if (imageBaseDir != null && !imageBaseDir.isEmpty()) {
                    if (bucket != null && !bucket.isEmpty()) {
                        filePath = Paths.get(imageBaseDir, bucket, path);
                    } else {
                        filePath = Paths.get(imageBaseDir, path);
                    }
                } else {
                    filePath = Paths.get(path);
                }
            }
            return Files.readAllBytes(filePath);

        } catch (IOException e) {
            throw new RuntimeException("Cannot load image from path=" + path
                    + ", bucket=" + bucket, e);
        }
    }

    // ===================== TABLE DATA (generic list) =====================

    private void processTableData(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {
        MainDocumentPart main = pkg.getMainDocumentPart();
        Body body = main.getContents().getBody();

        for (Object bodyObj : new ArrayList<>(body.getContent())) {
            Object u = XmlUtils.unwrap(bodyObj);
            if (!(u instanceof Tbl)) continue;

            Tbl tbl = (Tbl) u;
            handleTable(tbl, root);
        }
    }

    private void handleTable(Tbl tbl, Map<String, Object> root) {
        List<Tr> rows = new ArrayList<>();
        for (Object rObj : tbl.getContent()) {
            rows.add((Tr) XmlUtils.unwrap(rObj));
        }

        for (Tr row : new ArrayList<>(rows)) {
            String rowText = getRowText(row);
            Matcher m = LIST_IN_ROW.matcher(rowText);
            if (m.find()) {
                String listKey = m.group(1); // students, orders, ...
                fillTableForList(tbl, row, listKey, root);
            }
        }
    }

    @SuppressWarnings("unchecked")
    private void fillTableForList(Tbl tbl, Tr templateRow, String listKey,
                                  Map<String, Object> root) {
        Object value = resolveKey(root, listKey);
        if (!(value instanceof List)) {
            tbl.getContent().remove(templateRow);
            return;
        }

        List<?> list = (List<?>) value;
        int insertIndex = tbl.getContent().indexOf(templateRow);
        tbl.getContent().remove(templateRow);

        int index = 1;
        for (Object item : list) {
            Map<String, Object> itemCtx = toMap(item);
            itemCtx.put("index", index++); // dùng {{listKey.index}} nếu muốn

            Tr newRow = XmlUtils.deepCopy(templateRow);
            replaceRowScalars(newRow, listKey, itemCtx);
            tbl.getContent().add(insertIndex++, newRow);
        }
    }

    @SuppressWarnings("unchecked")
    private Map<String, Object> toMap(Object item) {
        if (item instanceof Map) {
            return (Map<String, Object>) item;
        }
        // POJO -> Map
        return new ObjectMapper().convertValue(item,
                new TypeReference<Map<String, Object>>() {});
    }

    private String getRowText(Tr row) {
        StringBuilder sb = new StringBuilder();
        for (Object tcObj : row.getContent()) {
            Tc cell = (Tc) XmlUtils.unwrap(tcObj);
            for (Object pObj : cell.getContent()) {
                P p = (P) XmlUtils.unwrap(pObj);
                sb.append(getParagraphText(p));
            }
        }
        return sb.toString();
    }

    private void replaceRowScalars(Tr row, String listKey, Map<String, Object> itemCtx) {
        for (Object tcObj : row.getContent()) {
            Tc cell = (Tc) XmlUtils.unwrap(tcObj);
            for (Object pObj : cell.getContent()) {
                P p = (P) XmlUtils.unwrap(pObj);
                String txt = getParagraphText(p);

                Matcher m = SCALAR.matcher(txt);
                StringBuffer sb = new StringBuffer();
                while (m.find()) {
                    String key = m.group(1).trim(); // students.name
                    String field = key;
                    if (key.startsWith(listKey + ".")) {
                        field = key.substring((listKey + ".").length());
                    }
                    Object val = itemCtx.get(field);
                    m.appendReplacement(sb, Matcher.quoteReplacement(
                            val != null ? String.valueOf(val) : ""));
                }
                m.appendTail(sb);
                setParagraphText(p, sb.toString());
            }
        }
    }

    // ===================== SUBJECTS (format dạng text) =====================

    @SuppressWarnings("unchecked")
    private void fillSubjectsText(Body body, Map<String, Object> root) {
        Object value = resolveKey(root, "subjects");
        if (!(value instanceof List)) return;

        List<Map<String, Object>> subjects = (List<Map<String, Object>>) value;
        List<Object> content = body.getContent();

        int dsIndex = -1;
        for (int i = 0; i < content.size(); i++) {
            Object u = XmlUtils.unwrap(content.get(i));
            if (u instanceof P) {
                String txt = getParagraphText((P) u).trim();
                if (txt.startsWith("Danh sách môn học")) {
                    dsIndex = i;
                    break;
                }
            }
        }
        if (dsIndex == -1) return;

        int insertIndex = dsIndex + 1;

        int maxRemove = Math.min(5, content.size() - insertIndex);
        for (int i = 0; i < maxRemove; i++) {
            content.remove(insertIndex);
        }

        for (Map<String, Object> sub : subjects) {
            String name   = sub.get("name")   != null ? String.valueOf(sub.get("name"))   : "";
            String credit = sub.get("credit") != null ? String.valueOf(sub.get("credit")) : "";
            String score  = sub.get("score")  != null ? String.valueOf(sub.get("score"))  : "";

            content.add(insertIndex++, createPlainParagraph("- " + name));
            content.add(insertIndex++, createPlainParagraph("| Tín chỉ: " + credit));
            content.add(insertIndex++, createPlainParagraph("| Điểm: " + score));
            content.add(insertIndex++, createPlainParagraph("")); // dòng trống giữa môn
        }
    }

    private P createPlainParagraph(String text) {
        P p = new P();
        R r = new R();
        Text t = new Text();
        t.setValue(text);
        r.getContent().add(t);
        p.getContent().add(r);
        return p;
    }

    // ===================== TEXT / BLOCK =====================

    private void processTextBlocks(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {
        List<Object> paragraphs;
        try {
            paragraphs = pkg.getMainDocumentPart()
                    .getJAXBNodesViaXPath("//w:p", true);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        boolean insideBlock = false;
        String currentKey = null;
        List<P> buffer = new ArrayList<>();

        for (Object obj : paragraphs) {
            P p = (P) XmlUtils.unwrap(obj);
            String text = getParagraphText(p);
            Matcher startM = BLOCK_START.matcher(text);
            Matcher endM = BLOCK_END.matcher(text);

            if (!insideBlock && startM.find()) {
                insideBlock = true;
                currentKey = startM.group(1).trim();
                buffer.clear();
                buffer.add(p);
                continue;
            }

            if (insideBlock) {
                buffer.add(p);
                if (endM.find()) {
                    boolean show = isTruthy(resolveKey(root, currentKey));
                    if (show) {
                        for (P blockP : buffer) {
                            String t = getParagraphText(blockP);
                            t = t.replace("{{?" + currentKey + "}}", "")
                                    .replace("{{/" + currentKey + "}}", "");
                            t = replaceScalars(t, root, currentKey);

                            if (t == null || t.trim().isEmpty()) {
                                deleteParagraph(pkg, blockP);
                            } else {
                                setParagraphText(blockP, t);
                            }
                        }
                    } else {
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

            // ngoài block: scalar thường
            String replaced = replaceScalars(text, root, null);
            setParagraphText(p, replaced);
        }
    }

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

    private void setParagraphText(P p, String newText) {
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

    private boolean isTruthy(Object value) {
        if (value == null) return false;
        if (value instanceof Boolean) return (Boolean) value;
        if (value instanceof String) return !((String) value).trim().isEmpty();
        if (value instanceof Collection) return !((Collection<?>) value).isEmpty();
        if (value instanceof Map) return !((Map<?, ?>) value).isEmpty();
        return true;
    }

    private Object resolveKey(Map<String, Object> root, String key) {
        if (key == null || key.isEmpty()) return null;
        String[] parts = key.split("\\.");
        Object current = root;
        for (String part : parts) {
            if (!(current instanceof Map)) return null;
            current = ((Map<?, ?>) current).get(part);
            if (current == null) return null;
        }
        return current;
    }

    private String replaceScalars(String text, Map<String, Object> root, String blockKey) {
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
            m.appendReplacement(sb, Matcher.quoteReplacement(
                    value != null ? String.valueOf(value) : ""));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    // ===================== CLEANUP PARAGRAPHS =====================

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
            String txt = getParagraphText(p).trim();
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
}

