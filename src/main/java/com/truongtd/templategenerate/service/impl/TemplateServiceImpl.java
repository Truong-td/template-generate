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
import java.math.BigInteger;
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
    private static final Pattern LIST_BLOCK_START =
            Pattern.compile("\\{\\{([a-zA-Z0-9_]+)}}");
    // paragraph chỉ chứa 1 scalar: {{avatar}}, {{user.avatar}}, ...
    private static final Pattern IMAGE_ONLY_PLACEHOLDER =
            Pattern.compile("\\{\\{([^{}]+)}}");

    private static final Pattern COND_START = Pattern.compile("\\{\\{\\?([^}]+)}}");
    private static final Pattern COND_END   = Pattern.compile("\\{\\{\\/([^}]+)}}");

    @Override
    public byte[] generateDocx(GenerateTemplateRequest request) {
        try {
            TemplateDataDto data = JsonUtils.parse(request);

            WordprocessingMLPackage pkg = WordprocessingMLPackage.load(
                    getClass().getResourceAsStream("/templates/template-report.docx"));

            Map<String, Object> context = buildRootContext(data);
            System.out.println("context = " + context);

            // 1. FlexData
            processFlexData(pkg, data.getFlexData());

            // 1) condition trước để xóa được cả table
            processConditionalBlocks(pkg, context);

            // 2. List-block cho table + text
            processListBlocks(pkg, context);

            // 3. Block điều kiện + scalar còn lại
            processTextBlocks(pkg, context);

            // 4. Dọn paragraph trống
            cleanupEmptyParagraphs(pkg);

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

            String s = Optional.ofNullable(getParagraphText((P) u)).orElse("").trim();
            Matcher ms = COND_START.matcher(s);
            if (!ms.matches()) { i++; continue; }

            String key = ms.group(1).trim().replaceAll("\\s+", ""); // remove spaces in key

            int end = -1;
            for (int j = i + 1; j < c.size(); j++) {
                Object uj = XmlUtils.unwrap(c.get(j));
                if (uj instanceof P) {
                    String ej = Optional.ofNullable(getParagraphText((P) uj)).orElse("").trim();
                    Matcher me = COND_END.matcher(ej);
                    if (me.matches()) {
                        String endKey = me.group(1).trim().replaceAll("\\s+", "");
                        if (endKey.equals(key)) { end = j; break; }
                    }
                }
            }
            if (end == -1) { i++; continue; }

            Object condVal = resolveKey(root, key);
            boolean show = isTruthy(condVal);

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

    private void applyVerticalMergeContinueAndClear(Tr row) {
        for (Object cellObj : row.getContent()) {
            Object cu = XmlUtils.unwrap(cellObj);
            if (!(cu instanceof Tc tc)) continue;

            TcPr pr = tc.getTcPr();
            if (pr == null) continue;

            TcPrInner.VMerge vm = pr.getVMerge();
            if (vm == null) continue;

            // các row sau phải là CONTINUE + clear text
            vm.setVal("continue");
            tc.getContent().clear();
            tc.getContent().add(new P());
        }
    }

    private void clearNonListCellsForSubsequentRow(Tr row, String listKey) {
        if (row == null) return;

        String marker = "{{" + listKey + ".";
        for (Object cellObj : row.getContent()) {
            Object cu = XmlUtils.unwrap(cellObj);
            if (!(cu instanceof Tc)) continue;
            Tc tc = (Tc) cu;

            String cellText = getTcText(tc);
            if (cellText == null) cellText = "";

            // Nếu cell KHÔNG chứa {{criteriaList.xxx}} thì coi là "tĩnh" -> clear
            if (!cellText.contains(marker)) {
                // clear toàn bộ content của cell
                tc.getContent().clear();
                tc.getContent().add(new P());
            }
        }
    }

    private String getTcText(Tc tc) {
        StringBuilder sb = new StringBuilder();
        for (Object o : tc.getContent()) {
            Object u = XmlUtils.unwrap(o);
            if (u instanceof P) {
                String t = getParagraphText((P) u);
                if (t != null) sb.append(t);
            }
        }
        return sb.toString();
    }

    private boolean handleImagePlaceholder(WordprocessingMLPackage pkg,
                                           P paragraph,
                                           Map<String, Object> ctx) throws Docx4JException {
        String txt = getParagraphText(paragraph);
        if (txt == null) return false;

        txt = txt.trim();
        Matcher m = IMAGE_ONLY_PLACEHOLDER.matcher(txt);
        if (!m.matches()) {
            return false; // paragraph không phải dạng "{{key}}" duy nhất
        }

        String key = m.group(1).trim(); // avatar, user.avatar, ...
        Object val = resolveKey(ctx, key);
        if (val == null) {
            // không có dữ liệu => xoá paragraph
            deleteParagraph(pkg, paragraph);
            return true;
        }

        String path = String.valueOf(val);
        // nếu không giống đường dẫn ảnh -> coi như text bình thường
        if (!looksLikeImagePath(path)) {
            return false;
        }

        try {
            Map<String, Object> imageData = new HashMap<>();
            imageData.put("bucket", null);
            imageData.put("path", path);

            Map<String, Object> block = new HashMap<>();
            block.put("imageData", imageData);

            P imgP = createFlexImageParagraph(pkg, block);

            Body body = pkg.getMainDocumentPart().getContents().getBody();
            List<Object> content = body.getContent();
            int idx = content.indexOf(paragraph);
            if (idx >= 0) {
                content.set(idx, imgP);
            } else {
                content.add(imgP);
                deleteParagraph(pkg, paragraph);
            }
            return true;
        } catch (Exception e) {
            throw new RuntimeException("Error inserting image for key=" + key
                    + ", path=" + path, e);
        }
    }

    // simple heuristic: path nhìn giống ảnh (URL, classpath, file ảnh)
    private boolean looksLikeImagePath(String path) {
        String p = path.toLowerCase(Locale.ROOT).trim();
        if (p.startsWith("http://") || p.startsWith("https://") || p.startsWith("classpath:")) {
            return true;
        }
        return p.endsWith(".png") || p.endsWith(".jpg") || p.endsWith(".jpeg")
                || p.endsWith(".gif") || p.endsWith(".bmp") || p.endsWith(".webp");
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
    @SuppressWarnings("unchecked")
    private void processListBlocks(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {
        MainDocumentPart main = pkg.getMainDocumentPart();
        Body body = main.getContents().getBody();
        List<Object> content = body.getContent();

        int i = 0;
        while (i < content.size()) {
            Object u = XmlUtils.unwrap(content.get(i));
            if (!(u instanceof P)) {
                i++;
                continue;
            }

            P pStart = (P) u;
            String txt = getParagraphText(pStart);
            if (txt == null) {
                i++;
                continue;
            }
            txt = txt.trim();

            Matcher m = LIST_BLOCK_START.matcher(txt);
            if (!m.matches()) {
                i++;
                continue;
            }

            String listKey = m.group(1); // students, subjects, ...

            Object listObj = resolveKey(root, listKey);
            if (!(listObj instanceof List)) {
                i++;
                continue;
            }
            List<?> list = (List<?>) listObj;
            if (list.isEmpty()) {
                // xoá luôn block nếu list rỗng
                // tìm endIndex rồi xoá
                int endIdx = findListBlockEndIndex(content, i, listKey);
                if (endIdx != -1) {
                    content.subList(i, endIdx + 1).clear();
                } else {
                    i++;
                }
                continue;
            }

            int endIndex = findListBlockEndIndex(content, i, listKey);
            if (endIndex == -1) {
                i++;
                continue;
            }

            // các node giữa {{listKey}} và {{/listKey}}
            List<Object> templateNodes = new ArrayList<>();
            for (int k = i + 1; k < endIndex; k++) {
                templateNodes.add(XmlUtils.deepCopy(content.get(k)));
            }

            // xây output nodes cho block này
            List<Object> outputNodes = buildListBlockOutput(listKey, list, templateNodes, root);

            // xoá block gốc và chèn output
            content.subList(i, endIndex + 1).clear();
            content.addAll(i, outputNodes);
            i += outputNodes.size();
        }
    }

    private int findListBlockEndIndex(List<Object> content, int startIndex, String listKey) {
        for (int j = startIndex + 1; j < content.size(); j++) {
            Object u2 = XmlUtils.unwrap(content.get(j));
            if (u2 instanceof P) {
                P pEnd = (P) u2;
                String t2 = getParagraphText(pEnd);
                if (t2 != null && t2.trim().equals("{{/" + listKey + "}}")) {
                    return j;
                }
            }
        }
        return -1;
    }
//    @SuppressWarnings("unchecked")
//    private List<Object> buildListBlockOutput(String listKey,
//                                              List<?> list,
//                                              List<Object> templateNodes,
//                                              Map<String, Object> root) {
//        List<Object> out = new ArrayList<>();
//
//        List<Tbl> tables = new ArrayList<>();
//        List<Object> nonTableNodes = new ArrayList<>();
//
//        for (Object node : templateNodes) {
//            Object u = XmlUtils.unwrap(node);
//            if (u instanceof Tbl) {
//                tables.add((Tbl) u);
//            } else if (u instanceof P && isParagraphEmpty((P) u)) {
//                // paragraph rỗng thì bỏ qua, ko tính là non-table
//                continue;
//            } else {
//                nonTableNodes.add(node);
//            }
//        }
//
//        // ===== CASE 1: block chỉ có đúng 1 TABLE (students) -> TABLE MODE =====
//        if (tables.size() == 1 && nonTableNodes.isEmpty()) {
//            Tbl templateTbl = tables.get(0);
//
//            // clone cả bảng làm bảng output
//            Tbl outTbl = XmlUtils.deepCopy(templateTbl);
//            out.add(outTbl);
//
//            // lấy list row trong bảng output (để replace cho item đầu tiên)
//            List<Tr> outRows = new ArrayList<>();
//            for (Object rObj : outTbl.getContent()) {
//                outRows.add((Tr) XmlUtils.unwrap(rObj));
//            }
//
//            // xác định các row nào là "template row" (có placeholder {{...}})
//            List<Integer> templateRowIdx = new ArrayList<>();
//            for (int idx = 0; idx < outRows.size(); idx++) {
//                String rowText = getRowText(outRows.get(idx));
//                if (rowText != null && rowText.contains("{{")) {
//                    templateRowIdx.add(idx);
//                }
//            }
//
//            // lấy bản gốc template row từ templateTbl (để dùng cho item 2,3,...)
//            List<Tr> originalTemplateRows = new ArrayList<>();
//            for (int idx : templateRowIdx) {
//                Tr r = (Tr) XmlUtils.unwrap(templateTbl.getContent().get(idx));
//                originalTemplateRows.add(r);
//            }
//
//            // --- item thứ 1: replace trực tiếp lên row trong outTbl ---
//            Map<String, Object> ctx1 = buildItemContext(listKey, list.get(0), 1, root);
//            for (int idx : templateRowIdx) {
//                Tr row = outRows.get(idx);
//                replaceScalarsDeep(row, ctx1);
//            }
//
//            // --- item thứ 2 trở đi: append thêm row vào cùng bảng ---
//            int insertPos = outTbl.getContent().size();
//            for (int itemIndex = 1; itemIndex < list.size(); itemIndex++) {
//                Map<String, Object> ctx = buildItemContext(listKey, list.get(itemIndex),
//                        itemIndex + 1, root);
//                for (Tr templRow : originalTemplateRows) {
//                    Tr newRow = XmlUtils.deepCopy(templRow);
//                    applyVerticalMergeContinueAndClear(newRow);
//                    replaceScalarsDeep(newRow, ctx);
//                    outTbl.getContent().add(insertPos++, newRow);
//                }
//            }
//
//            return out;
//        }
//
//        // ===== CASE 2: TEXT MODE (subjects, orders, ...) =====
//        // -> lặp nguyên block cho mỗi item (behavior cũ)
//        for (int itemIndex = 0; itemIndex < list.size(); itemIndex++) {
//            Map<String, Object> ctx = buildItemContext(listKey, list.get(itemIndex),
//                    itemIndex + 1, root);
//
//            for (Object tplNode : templateNodes) {
//                Object copy = XmlUtils.deepCopy(tplNode);
//                replaceScalarsDeep(copy, ctx);
//                out.add(copy);
//            }
//        }
//
//        return out;
//    }
    @SuppressWarnings("unchecked")
    private List<Object> buildListBlockOutput(String listKey,
                                              List<?> list,
                                              List<Object> templateNodes,
                                              Map<String, Object> root) {
        List<Object> out = new ArrayList<>();

        // tách table nodes và non-table nodes
        List<Tbl> tables = new ArrayList<>();
        List<Object> nonTableNodes = new ArrayList<>();

        for (Object node : templateNodes) {
            Object u = XmlUtils.unwrap(node);
            if (u instanceof Tbl) {
                tables.add((Tbl) u);
            } else if (u instanceof P && isParagraphEmpty((P) u)) {
                // bỏ paragraph trống
                continue;
            } else {
                nonTableNodes.add(node);
            }
        }

        // ===================== TABLE MODE =====================
        // block chỉ có đúng 1 table và không có node khác (ngoài paragraph trống)
        if (tables.size() == 1 && nonTableNodes.isEmpty()) {
            Tbl templateTbl = tables.get(0);

            // bảng output (giữ header 1 lần)
            Tbl outTbl = XmlUtils.deepCopy(templateTbl);
            out.add(outTbl);

            // rows của bảng output (để replace cho item #1)
            List<Tr> outRows = new ArrayList<>();
            for (Object rObj : outTbl.getContent()) {
                outRows.add((Tr) XmlUtils.unwrap(rObj));
            }

            // xác định những row là "template row" (row có placeholder {{...}})
            List<Integer> templateRowIdx = new ArrayList<>();
            for (int idx = 0; idx < outRows.size(); idx++) {
                String rowText = getRowText(outRows.get(idx));
                if (rowText != null && rowText.contains("{{")) {
                    templateRowIdx.add(idx);
                }
            }

            // nếu không có template row => coi như không làm gì (tránh lỗi)
            if (templateRowIdx.isEmpty()) {
                return out;
            }

            // lấy bản gốc template row từ templateTbl để dùng clone cho item #2+
            List<Tr> originalTemplateRows = new ArrayList<>();
            for (int idx : templateRowIdx) {
                Object srcRowObj = templateTbl.getContent().get(idx);
                Tr srcRow = (Tr) XmlUtils.unwrap(srcRowObj);
                originalTemplateRows.add(srcRow);
            }

            // -------- item #1: replace trực tiếp trên outTbl --------
            Map<String, Object> ctx1 = buildItemContext(listKey, list.get(0), 1, root);
            for (int idx : templateRowIdx) {
                Tr row = outRows.get(idx);
                // replace cả row (bao gồm cell)
                replaceScalarsDeep(row, ctx1);
            }

            // -------- item #2+: clone template rows và append vào outTbl --------
            int insertPos = outTbl.getContent().size();
            for (int itemIndex = 1; itemIndex < list.size(); itemIndex++) {
                Map<String, Object> ctx = buildItemContext(listKey, list.get(itemIndex), itemIndex + 1, root);

                for (Tr templRow : originalTemplateRows) {
                    Tr newRow = XmlUtils.deepCopy(templRow);

                    // ✅ MERGE FIX: nếu row có vMerge thì chuyển sang continue và clear text trong cell merge
                    applyVerticalMergeContinueAndClear(newRow);

                    // ✅ fix lặp text kiểu "CTTD: test" nếu template không merge
                    clearNonListCellsForSubsequentRow(newRow, listKey);

                    replaceScalarsDeep(newRow, ctx);

                    // ✅ nếu row sau cùng bị rỗng => bỏ luôn, không add vào table
                    if (isRowBlank(newRow)) {
                        continue;
                    }
                    outTbl.getContent().add(insertPos++, newRow);
                }
            }

            return out;
        }

        // ===================== TEXT MODE =====================
        // lặp nguyên block cho mỗi item (subjects, orders dạng text, hoặc block có nhiều node)
        for (int itemIndex = 0; itemIndex < list.size(); itemIndex++) {
            Map<String, Object> ctx = buildItemContext(listKey, list.get(itemIndex), itemIndex + 1, root);

            for (Object tplNode : templateNodes) {
                Object copy = XmlUtils.deepCopy(tplNode);
                replaceScalarsDeep(copy, ctx);
                out.add(copy);
            }
        }

        return out;
    }
    private boolean isRowBlank(Tr row) {
        if (row == null) return true;

        for (Object cellObj : row.getContent()) {
            Object cu = XmlUtils.unwrap(cellObj);
            if (!(cu instanceof Tc)) continue;
            Tc tc = (Tc) cu;

            String cellText = getTcText(tc);
            if (cellText != null && !cellText.trim().isEmpty()) {
                return false;
            }
            // nếu cell có drawing (ảnh) cũng coi là không rỗng
            if (containsDrawing(tc)) {
                return false;
            }
        }
        return true;
    }
    private boolean containsDrawing(Tc tc) {
        for (Object o : tc.getContent()) {
            Object u = XmlUtils.unwrap(o);
            if (u instanceof P) {
                P p = (P) u;
                for (Object rObj : p.getContent()) {
                    Object ru = XmlUtils.unwrap(rObj);
                    if (ru instanceof R) {
                        for (Object rc : ((R) ru).getContent()) {
                            Object cu = XmlUtils.unwrap(rc);
                            if (cu instanceof Drawing) return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    @SuppressWarnings("unchecked")
    private Map<String, Object> buildItemContext(String listKey,
                                                 Object itemObj,
                                                 int index,
                                                 Map<String, Object> root) {
        Map<String, Object> itemMap;
        if (itemObj instanceof Map) {
            itemMap = (Map<String, Object>) itemObj;
        } else {
            itemMap = new ObjectMapper().convertValue(itemObj, new TypeReference<Map<String, Object>>() {});
        }

        Map<String, Object> itemWithIndex = new HashMap<>(itemMap);
        itemWithIndex.put("index", index);

        Map<String, Object> ctx = new HashMap<>(root);
        ctx.put(listKey, itemWithIndex); // {{students.name}}
        ctx.putAll(itemWithIndex);       // {{name}}, {{age}}, {{index}}
        return ctx;
    }

    private String getRowText(Tr row) {
        StringBuilder sb = new StringBuilder();
        for (Object cellObj : row.getContent()) {
            Object cu = XmlUtils.unwrap(cellObj);
            if (cu instanceof Tc) {
                Tc tc = (Tc) cu;
                for (Object pObj : tc.getContent()) {
                    Object pu = XmlUtils.unwrap(pObj);
                    if (pu instanceof P) {
                        String t = getParagraphText((P) pu);
                        if (t != null) sb.append(t);
                    }
                }
            }
        }
        return sb.toString();
    }

    private boolean isParagraphEmpty(P p) {
        String t = getParagraphText(p);
        return t == null || t.trim().isEmpty();
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
// ==================== FLEX DATA ====================

    @SuppressWarnings("unchecked")
    private void processFlexData(WordprocessingMLPackage pkg,
                                 Map<String, Object> flexData) throws Exception {
        if (flexData == null || flexData.isEmpty()) return;

        for (Map.Entry<String, Object> e : flexData.entrySet()) {
            String key = e.getKey();      // ví dụ: note, historyFormation,...
            Object val = e.getValue();

            List<Object> blocks = null;

            if (val instanceof List) {
                blocks = (List<Object>) val;
            } else if (val instanceof Map) {
                blocks = convertSimpleFlexObjectToBlocks((Map<String, Object>) val);
            } else if (val instanceof String) {
                // chỉ text
                Map<String, Object> block = new HashMap<>();
                block.put("type", "text");
                Map<String, Object> textData = new HashMap<>();
                textData.put("data", val);
                textData.put("isBold", false);
                block.put("textData", textData);
                blocks = Collections.singletonList(block);
            }

            if (blocks == null || blocks.isEmpty()) continue;
            processOneFlexBlock(pkg, key, blocks);
        }
    }

    /**
     * Cho JSON đơn giản dạng:
     * { "text": "...", "table":[["A","B"],[1,2]], "image":"http://..." }
     */
    @SuppressWarnings("unchecked")
    private List<Object> convertSimpleFlexObjectToBlocks(Map<String, Object> obj) {
        List<Object> blocks = new ArrayList<>();

        // text
        Object textVal = obj.get("text");
        if (textVal != null) {
            Map<String, Object> block = new HashMap<>();
            block.put("type", "text");
            Map<String, Object> textData = new HashMap<>();
            textData.put("data", String.valueOf(textVal));
            textData.put("isBold", false);
            block.put("textData", textData);
            blocks.add(block);
        }

        // table [[...],[...]]
        Object tableVal = obj.get("table");
        if (tableVal instanceof List) {
            List<List<?>> raw = (List<List<?>>) tableVal;
            List<List<Map<String, Object>>> matrix = new ArrayList<>();
            for (int r = 0; r < raw.size(); r++) {
                List<?> row = raw.get(r);
                List<Map<String, Object>> rowOut = new ArrayList<>();
                for (Object cellVal : row) {
                    Map<String, Object> cell = new HashMap<>();
                    cell.put("data", cellVal != null ? String.valueOf(cellVal) : "");
                    cell.put("isBold", r == 0); // dòng đầu in đậm
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

        // image
        Object imgVal = obj.get("image");
        if (imgVal != null) {
            Map<String, Object> block = new HashMap<>();
            block.put("type", "image");
            Map<String, Object> imageData = new HashMap<>();
            imageData.put("bucket", null);
            imageData.put("path", String.valueOf(imgVal));
            block.put("imageData", imageData);
            blocks.add(block);
        }

        return blocks;
    }

    @SuppressWarnings("unchecked")
    private void processOneFlexBlock(WordprocessingMLPackage pkg,
                                     String placeholderKey,
                                     List<Object> blocks) throws Exception {
        MainDocumentPart main = pkg.getMainDocumentPart();
        Body body = main.getContents().getBody();

        String tag = "{{" + placeholderKey + "}}";

        P placeholderPara = null;
        for (Object o : body.getContent()) {
            Object u = XmlUtils.unwrap(o);
            if (u instanceof P) {
                P p = (P) u;
                String txt = getParagraphText(p);
                if (txt != null && txt.contains(tag)) {
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
                    body.getContent().add(insertIndex++, createFlexTextParagraph(block));
                    break;
                case "table":
                    body.getContent().add(insertIndex++, createFlexTable(block));
                    break;
                case "image":
                    body.getContent().add(insertIndex++, createFlexImageParagraph(pkg, block));
                    break;
                default:
                    break;
            }
        }
    }

    @SuppressWarnings("unchecked")
    private P createFlexTextParagraph(Map<String, Object> block) {
        Map<String, Object> textData = (Map<String, Object>) block.get("textData");
        String data = textData != null ? (String) textData.get("data") : "";
        boolean isBold = textData != null && Boolean.TRUE.equals(textData.get("isBold"));

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
    private Tbl createFlexTable(Map<String, Object> block) {
        Map<String, Object> tableData = (Map<String, Object>) block.get("tableData");
        List<List<Map<String, Object>>> matrix =
                (List<List<Map<String, Object>>>) tableData.get("data");

        Tbl tbl = new Tbl();

        TblPr tblPr = new TblPr();
        TblBorders borders = new TblBorders();
        borders.setTop(createBorder());
        borders.setBottom(createBorder());
        borders.setLeft(createBorder());
        borders.setRight(createBorder());
        borders.setInsideH(createBorder());
        borders.setInsideV(createBorder());
        tblPr.setTblBorders(borders);
        tbl.setTblPr(tblPr);

        for (List<Map<String, Object>> rowData : matrix) {
            Tr tr = new Tr();
            for (Map<String, Object> cellData : rowData) {
                String text = (String) cellData.get("data");
                boolean isBold = Boolean.TRUE.equals(cellData.get("isBold"));

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

    private CTBorder createBorder() {
        CTBorder border = new CTBorder();
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.ZERO);
        border.setColor("000000");
        return border;
    }

    @SuppressWarnings("unchecked")
    private P createFlexImageParagraph(WordprocessingMLPackage pkg,
                                       Map<String, Object> block) throws Exception {
        Map<String, Object> imageData = (Map<String, Object>) block.get("imageData");
        if (imageData == null) return new P();

        String bucket = (String) imageData.get("bucket");
        String path = (String) imageData.get("path");

        byte[] bytes = loadImageBytes(bucket, path);
        if (bytes == null || bytes.length == 0) {
            P p = new P();
            R r = new R();
            Text t = new Text();
            t.setValue("[IMAGE NOT FOUND: " + path + "]");
            r.getContent().add(t);
            p.getContent().add(r);
            return p;
        }

        BinaryPartAbstractImage imagePart =
                BinaryPartAbstractImage.createImagePart(pkg, bytes);

        int id1 = (int) (Math.random() * 10000);
        int id2 = (int) (Math.random() * 10000);

        Inline inline = imagePart.createImageInline(
                "flex-img", "flex-img", id1, id2, 6000, false);

        Drawing drawing = new Drawing();
        drawing.getAnchorOrInline().add(inline);

        R r = new R();
        r.getContent().add(drawing);

        P p = new P();
        p.getContent().add(r);
        return p;
    }

    private byte[] loadImageBytes(String bucket, String path) {
        if (path == null || path.trim().isEmpty()) return new byte[0];
        path = path.trim();

        try {
            if (path.startsWith("http://") || path.startsWith("https://")) {
                try (InputStream is = new URL(path).openStream()) {
                    return IOUtils.toByteArray(is);
                }
            }

            if (path.startsWith("classpath:")) {
                String cp = path.substring("classpath:".length());
                String res = cp.startsWith("/") ? cp : "/" + cp;
                try (InputStream is = getClass().getResourceAsStream(res)) {
                    if (is == null) throw new FileNotFoundException("Classpath not found: " + cp);
                    return IOUtils.toByteArray(is);
                }
            }

            Path filePath;
            if (Paths.get(path).isAbsolute()) {
                filePath = Paths.get(path);
            } else if (imageBaseDir != null && !imageBaseDir.isEmpty()) {
                if (bucket != null && !bucket.isEmpty()) {
                    filePath = Paths.get(imageBaseDir, bucket, path);
                } else {
                    filePath = Paths.get(imageBaseDir, path);
                }
            } else {
                filePath = Paths.get(path);
            }
            return Files.readAllBytes(filePath);
        } catch (IOException e) {
            throw new RuntimeException("Cannot load image from path=" + path
                    + ", bucket=" + bucket, e);
        }
    }

//    @SuppressWarnings("unchecked")
//    private Map<String, Object> toMap(Object item) {
//        if (item instanceof Map) return (Map<String, Object>) item;
//        return MAPPER.convertValue(item, new TypeReference<Map<String, Object>>() {});
//    }

    private void replaceScalarsDeep(Object node, Map<String, Object> ctx) {
        Object u = XmlUtils.unwrap(node);
        if (u instanceof P) {
            P p = (P) u;
            String text = getParagraphText(p);
            if (text != null && text.contains("{{")) {
                String replaced = replaceScalars(text, ctx);
                setParagraphText(p, replaced);
            }
        } else if (u instanceof ContentAccessor) {
            List<Object> children = ((ContentAccessor) u).getContent();
            for (Object child : children) {
                replaceScalarsDeep(child, ctx);
            }
        }
    }

    // ==================== TEXT BLOCKS (conditional + scalar) ====================

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
            if (text == null) text = "";

            Matcher startM = BLOCK_START.matcher(text);
            Matcher endM = BLOCK_END.matcher(text);

            // ---- BẮT ĐẦU BLOCK: {{?key}} ----
            if (!insideBlock && startM.find()) {
                insideBlock = true;
                currentKey = startM.group(1).trim(); // vd: "user"
                buffer.clear();
                buffer.add(p);
                continue;
            }

            if (insideBlock) {
                buffer.add(p);

                // ---- KẾT THÚC BLOCK: {{/key}} ----
                if (endM.find()) {
                    Object blockObj = resolveKey(root, currentKey);
                    boolean show = isTruthy(blockObj);

                    if (show) {
                        Map<String, Object> ctxForBlock = new HashMap<>(root);
                        if (blockObj instanceof Map) {
                            @SuppressWarnings("unchecked")
                            Map<String, Object> sub = (Map<String, Object>) blockObj;
                            ctxForBlock.putAll(sub);
                        }

                        for (P blockP : buffer) {
                            String t = getParagraphText(blockP);
                            if (t == null) t = "";

                            // bỏ marker điều kiện
                            t = t.replace("{{?" + currentKey + "}}", "")
                                    .replace("{{/" + currentKey + "}}", "");
                            setParagraphText(blockP, t);

                            // 1) Nếu paragraph là placeholder ảnh duy nhất -> chèn ảnh
                            if (handleImagePlaceholder(pkg, blockP, ctxForBlock)) {
                                continue;
                            }

                            // 2) Còn lại xử lý scalar text như cũ
                            String textAfter = getParagraphText(blockP);
                            if (textAfter != null && textAfter.contains("{{")) {
                                textAfter = replaceScalars(textAfter, ctxForBlock);
                                if (textAfter.trim().isEmpty()) {
                                    deleteParagraph(pkg, blockP);
                                } else {
                                    setParagraphText(blockP, textAfter);
                                }
                            } else if (textAfter == null || textAfter.trim().isEmpty()) {
                                deleteParagraph(pkg, blockP);
                            }
                        }
                    } else {
                        // không thỏa điều kiện -> xoá toàn bộ block
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

            // ---- ngoài mọi block điều kiện ----
            if (handleImagePlaceholder(pkg, p, root)) {
                continue;
            }
            // chỉ xử lý scalar nếu paragraph còn chứa {{ }}
            if (!text.contains("{{")) continue;

            String replaced = replaceScalars(text, root);
            setParagraphText(p, replaced);
        }
    }

    private String replaceScalars(String text, Map<String, Object> ctx) {
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

    // ==================== HELPERS ====================

    private String getParagraphText(P p) {
        StringBuilder sb = new StringBuilder();
        for (Object o : p.getContent()) {
            Object u = XmlUtils.unwrap(o);
            if (u instanceof R) {
                for (Object rc : ((R) u).getContent()) {
                    Object cu = XmlUtils.unwrap(rc);
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

    private boolean isTruthy(Object v) {
        if (v == null) return false;
        if (v instanceof Boolean) return (Boolean) v;
        if (v instanceof String s) {
            String x = s.trim();
            if (x.isEmpty()) return false;
            if ("false".equalsIgnoreCase(x)) return false;
            if ("true".equalsIgnoreCase(x)) return true;
            return true;
        }
        if (v instanceof Collection<?> c) return !c.isEmpty();
        if (v instanceof Map<?, ?> m) return !m.isEmpty();
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
            String txt = getParagraphText(p);
            boolean isEmpty = txt == null || txt.trim().isEmpty();

            if (isEmpty) {
                if (prevEmpty) {
                    it.remove();
                } else {
                    prevEmpty = true;
                }
            } else {
                prevEmpty = false;
            }
        }
    }

}

