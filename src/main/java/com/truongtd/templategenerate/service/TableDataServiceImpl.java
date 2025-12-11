//package com.truongtd.templategenerate.service;
//
//import com.fasterxml.jackson.core.type.TypeReference;
//import com.fasterxml.jackson.databind.ObjectMapper;
//import org.docx4j.XmlUtils;
//import org.docx4j.openpackaging.exceptions.Docx4JException;
//import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
//import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
//import org.docx4j.wml.Body;
//import org.docx4j.wml.P;
//import org.docx4j.wml.R;
//import org.docx4j.wml.Tbl;
//import org.docx4j.wml.Tc;
//import org.docx4j.wml.Text;
//import org.docx4j.wml.Tr;
//import org.springframework.stereotype.Service;
//
//import java.util.ArrayList;
//import java.util.List;
//import java.util.Map;
//import java.util.regex.Matcher;
//import java.util.regex.Pattern;
//
//import static com.msb.bpm.approval.templategenerator.service.impl.TemplateServiceImpl.SCALAR;
//
//@Service
//public class TableDataServiceImpl implements TableDataService {
//
//    private static final Pattern LIST_IN_ROW = Pattern.compile("\\{\\{([a-zA-Z0-9_]+)\\.[^}]+}}");
//
//    @Override
//    public void processTableData(WordprocessingMLPackage pkg, Map<String, Object> root) {
//        try {
//            MainDocumentPart main = pkg.getMainDocumentPart();
//            Body body = main.getContents().getBody();
//            // ==== 1. BẢNG SINH VIÊN (students) ====
//            Tbl studentsTbl = null;
//            for (Object bodyObj : body.getContent()) {
//                Object u = XmlUtils.unwrap(bodyObj);
//                if (u instanceof Tbl) {
//                    studentsTbl = (Tbl) u;
//                    handleTable(studentsTbl, root);
//                    break; // bảng đầu tiên
//                }
//            }
////            if (studentsTbl != null) {
////                fillStudentsTable(studentsTbl, root);
////            }
//
//            // ==== 2. DANH SÁCH MÔN HỌC (subjects) ====
//            fillSubjectsText(body, root);
//        } catch (Docx4JException e) {
//            throw new RuntimeException(e);
//        }
//    }
//
//    private void handleTable(Tbl tbl, Map<String, Object> root) {
//        List<Tr> rows = new ArrayList<>();
//        for (Object rObj : tbl.getContent()) {
//            rows.add((Tr) XmlUtils.unwrap(rObj));
//        }
//
//        // Bỏ qua hàng header đầu tiên nếu muốn
//        for (int i = 0; i < rows.size(); i++) {
//            Tr row = rows.get(i);
//            String rowText = getRowText(row);
//
//            Matcher m = LIST_IN_ROW.matcher(rowText);
//            if (m.find()) {
//                String listKey = m.group(1);   // ví dụ "students", "subjects", "orders"...
//                fillTableForList(tbl, row, listKey, root);
//            }
//        }
//    }
//
//    @SuppressWarnings("unchecked")
//    private void fillStudentsTable(Tbl tbl, Map<String, Object> root) {
//        Object value = resolveKey(root, "students");
//        if (!(value instanceof List)) {
//            // Không có data => xóa luôn row template nếu tồn tại
//            if (tbl.getContent().size() > 1) {
//                tbl.getContent().remove(1);
//            }
//            return;
//        }
//
//        List<Map<String, Object>> students = (List<Map<String, Object>>) value;
//        if (tbl.getContent().size() < 2) return; // phải có header + template
//
//        // row 0: header, row 1: template
//        Tr templateRow = (Tr) XmlUtils.unwrap(tbl.getContent().get(1));
//
//        // bỏ row template khỏi bảng trước
//        tbl.getContent().remove(1);
//
//        int insertIndex = 1;
//        int stt = 1;
//
//        for (Map<String, Object> student : students) {
//            Tr newRow = XmlUtils.deepCopy(templateRow);
//
//            // Giả định: 3 cột: STT | Name | Age
//            List<?> cellObjs = newRow.getContent();
//            if (cellObjs.size() >= 3) {
//                Tc sttCell = (Tc) XmlUtils.unwrap(cellObjs.get(0));
//                Tc nameCell = (Tc) XmlUtils.unwrap(cellObjs.get(1));
//                Tc ageCell  = (Tc) XmlUtils.unwrap(cellObjs.get(2));
//
//                setCellText(sttCell, String.valueOf(stt++));
//                setCellText(nameCell, student.get("name") != null ? String.valueOf(student.get("name")) : "");
//                setCellText(ageCell,  student.get("age")  != null ? String.valueOf(student.get("age"))  : "");
//            }
//
//            tbl.getContent().add(insertIndex++, newRow);
//        }
//    }
//
//    @SuppressWarnings("unchecked")
//    private void fillSubjectsText(Body body, Map<String, Object> root) {
//        Object value = resolveKey(root, "subjects");
//        if (!(value instanceof List)) return;
//
//        List<Map<String, Object>> subjects = (List<Map<String, Object>>) value;
//        List<Object> content = body.getContent();
//
//        // Tìm paragraph "Danh sách môn học:"
//        int dsIndex = -1;
//        for (int i = 0; i < content.size(); i++) {
//            Object u = XmlUtils.unwrap(content.get(i));
//            if (u instanceof P) {
//                String txt = getParagraphText((P) u).trim();
//                if (txt.startsWith("Danh sách môn học")) {
//                    dsIndex = i;
//                    break;
//                }
//            }
//        }
//        if (dsIndex == -1) return;
//
//        int insertIndex = dsIndex + 1;
//
//        // Xóa mấy paragraph cũ phía dưới (đang in List.toString + template cũ)
//        // Ở file hiện tại thường là 3–4 dòng, mình xoá tối đa 5 cho chắc
//        int maxRemove = Math.min(5, content.size() - insertIndex);
//        for (int i = 0; i < maxRemove; i++) {
//            content.remove(insertIndex);
//        }
//
//        // Chèn lại theo format:
//        // - subjectA
//        // | Tín chỉ: 00
//        // | Điểm: 10
//        for (Map<String, Object> sub : subjects) {
//            String name   = sub.get("name")   != null ? String.valueOf(sub.get("name"))   : "";
//            String credit = sub.get("credit") != null ? String.valueOf(sub.get("credit")) : "";
//            String score  = sub.get("score")  != null ? String.valueOf(sub.get("score"))  : "";
//
//            content.add(insertIndex++, createPlainParagraph("- " + name));
//            content.add(insertIndex++, createPlainParagraph("| Tín chỉ: " + credit));
//            content.add(insertIndex++, createPlainParagraph("| Điểm: " + score));
//            content.add(insertIndex++, createPlainParagraph("")); // dòng trống giữa các môn
//        }
//    }
//
//    private P createPlainParagraph(String text) {
//        P p = new P();
//        R r = new R();
//        Text t = new Text();
//        t.setValue(text);
//        r.getContent().add(t);
//        p.getContent().add(r);
//        return p;
//    }
//
//    private void setCellText(Tc cell, String text) {
//        cell.getContent().clear();
//        P p = new P();
//        R r = new R();
//        Text t = new Text();
//        t.setValue(text);
//        r.getContent().add(t);
//        p.getContent().add(r);
//        cell.getContent().add(p);
//    }
//
//    private String getRowText(Tr row) {
//        StringBuilder sb = new StringBuilder();
//        for (Object tcObj : row.getContent()) {
//            Tc cell = (Tc) XmlUtils.unwrap(tcObj);
//            for (Object pObj : cell.getContent()) {
//                P p = (P) XmlUtils.unwrap(pObj);
//                sb.append(getParagraphText(p));
//            }
//        }
//        return sb.toString();
//    }
//
//    @SuppressWarnings("unchecked")
//    private void fillTableForList(Tbl tbl, Tr templateRow, String listKey, Map<String, Object> root) {
//        Object value = resolveKey(root, listKey);
//        if (!(value instanceof List)) {
//            // Không có dữ liệu => xóa row template
//            tbl.getContent().remove(templateRow);
//            return;
//        }
//
//        List<?> list = (List<?>) value;
//
//        int insertIndex = tbl.getContent().indexOf(templateRow);
//        tbl.getContent().remove(templateRow);
//
//        int index = 1;
//        for (Object item : list) {
//            Map<String, Object> itemCtx = toMap(item); // convert POJO → Map nếu cần
//            itemCtx.put("index", index++); // dùng {{listKey.index}} nếu muốn
//
//            Tr newRow = XmlUtils.deepCopy(templateRow);
//            replaceRowScalars(newRow, listKey, itemCtx);
//            tbl.getContent().add(insertIndex++, newRow);
//        }
//    }
//
//    @SuppressWarnings("unchecked")
//    private Map<String, Object> toMap(Object item) {
//        if (item instanceof Map) return (Map<String, Object>) item;
//        // nếu là POJO, convert bằng ObjectMapper
//        return new ObjectMapper().convertValue(item, new TypeReference<Map<String,Object>>() {});
//    }
//
////    private void replaceRowScalars(Tr row, String listKey, Map<String, Object> itemCtx) {
////        for (Object tcObj : row.getContent()) {
////            Tc cell = (Tc) XmlUtils.unwrap(tcObj);
////            for (Object pObj : cell.getContent()) {
////                P p = (P) XmlUtils.unwrap(pObj);
////                String txt = getParagraphText(p);
////
////                // cho phép viết {{students.name}} hoặc {{students.index}}
////                Matcher m = SCALAR.matcher(txt);
////                StringBuffer sb = new StringBuffer();
////                while (m.find()) {
////                    String key = m.group(1).trim(); // ví dụ "students.name"
////                    String field = key;
////                    if (key.startsWith(listKey + ".")) {
////                        field = key.substring((listKey + ".").length());
////                    }
////                    Object value = itemCtx.get(field);
////                    m.appendReplacement(sb, Matcher.quoteReplacement(
////                            value != null ? String.valueOf(value) : ""));
////                }
////                m.appendTail(sb);
////                setParagraphText(p, sb.toString());
////            }
////        }
////    }
//    private void replaceRowScalars(Tr row, String listKey, Map<String, Object> itemCtx) {
//        for (Object tcObj : row.getContent()) {
//            Tc cell = (Tc) XmlUtils.unwrap(tcObj);
//            for (Object pObj : cell.getContent()) {
//                P p = (P) XmlUtils.unwrap(pObj);
//                String txt = getParagraphText(p);
//
//                Matcher m = SCALAR.matcher(txt); // {{...}}
//                StringBuffer sb = new StringBuffer();
//                while (m.find()) {
//                    String key = m.group(1).trim(); // vd "students.name"
//                    String field = key;
//
//                    if (key.startsWith(listKey + ".")) {
//                        field = key.substring((listKey + ".").length());
//                    }
//
//                    Object val = itemCtx.get(field);
//                    m.appendReplacement(sb, Matcher.quoteReplacement(
//                            val != null ? String.valueOf(val) : ""));
//                }
//                m.appendTail(sb);
//                setParagraphText(p, sb.toString());
//            }
//        }
//    }
//
//    private String getParagraphText(P p) {
//        StringBuilder sb = new StringBuilder();
//        for (Object o : p.getContent()) {
//            Object u = XmlUtils.unwrap(o);
//            if (u instanceof R) {
//                for (Object c : ((R) u).getContent()) {
//                    Object cu = XmlUtils.unwrap(c);
//                    if (cu instanceof Text) {
//                        sb.append(((Text) cu).getValue());
//                    }
//                }
//            }
//        }
//        return sb.toString();
//    }
//
//    private void setParagraphText(P p, String newText) {
//        // clear all run and replace bằng 1 run mới
//        p.getContent().clear();
//        R run = new R();
//        Text text = new Text();
//        text.setValue(newText);
//        run.getContent().add(text);
//        p.getContent().add(run);
//    }
//
//    private Object resolveKey(Map<String, Object> root, String key) {
//        // hỗ trợ nested: user.name, application.name...
//        String[] parts = key.split("\\.");
//        Object current = root;
//        for (String part : parts) {
//            if (!(current instanceof Map)) return null;
//            current = ((Map<?, ?>) current).get(part);
//            if (current == null) return null;
//        }
//        return current;
//    }
//}
