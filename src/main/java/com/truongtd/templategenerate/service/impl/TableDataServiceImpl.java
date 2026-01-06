package com.truongtd.templategenerate.service.impl;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.service.TableDataService;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

import static com.truongtd.templategenerate.util.StringUtils.LIST_BLOCK_START;

@Service
@Slf4j
public class TableDataServiceImpl implements TableDataService {
    private final DocxTemplateEngine templateEngine = new DocxTemplateEngine();

    @Override
    public void processTableData(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {
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
            String txt = templateEngine.getParagraphText(pStart);
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

            Object listObj = templateEngine.resolveKey(root, listKey);
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
                String t2 = templateEngine.getParagraphText(pEnd);
                if (t2 != null && t2.trim().equals("{{/" + listKey + "}}")) {
                    return j;
                }
            }
        }
        return -1;
    }
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

                    // MERGE FIX: nếu row có vMerge thì chuyển sang continue và clear text trong cell merge
                    templateEngine.applyVerticalMergeContinueAndClear(newRow);

                    // fix lặp text kiểu "CTTD: test" nếu template không merge
                    templateEngine.clearNonListCellsForSubsequentRow(newRow, listKey);

                    replaceScalarsDeep(newRow, ctx);

                    // nếu row sau cùng bị rỗng => bỏ luôn, không add vào table
                    if (templateEngine.isRowBlank(newRow)) {
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

    private void replaceScalarsDeep(Object node, Map<String, Object> ctx) {
        Object u = XmlUtils.unwrap(node);
        if (u instanceof P) {
            P p = (P) u;
            String text = templateEngine.getParagraphText(p);
            if (text != null && text.contains("{{")) {
                String replaced = templateEngine.replaceScalars(text, ctx);
                templateEngine.setParagraphText(p, replaced);
            }
        } else if (u instanceof ContentAccessor) {
            List<Object> children = ((ContentAccessor) u).getContent();
            for (Object child : children) {
                replaceScalarsDeep(child, ctx);
            }
        }
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
                        String t = templateEngine.getParagraphText((P) pu);
                        if (t != null) sb.append(t);
                    }
                }
            }
        }
        return sb.toString();
    }

    private boolean isParagraphEmpty(P p) {
        String t = templateEngine.getParagraphText(p);
        return t == null || t.trim().isEmpty();
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
}
