package com.truongtd.templategenerate.service.impl;

import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.service.TextDataService;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Body;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;

import static com.truongtd.templategenerate.util.StringUtils.BLOCK_END;
import static com.truongtd.templategenerate.util.StringUtils.BLOCK_START;
import static com.truongtd.templategenerate.util.StringUtils.IMAGE_ONLY_PLACEHOLDER;

@Service
public class TextDataServiceImpl implements TextDataService {
    private final DocxTemplateEngine templateEngine = new DocxTemplateEngine();
    @Override
    public void processTextBlocks(WordprocessingMLPackage pkg, Map<String, Object> root) throws Docx4JException {

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
            String text = templateEngine.getParagraphText(p);
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
                    Object blockObj = templateEngine.resolveKey(root, currentKey);
                    boolean show = templateEngine.isTruthy(blockObj);

                    if (show) {
                        Map<String, Object> ctxForBlock = new HashMap<>(root);
                        if (blockObj instanceof Map) {
                            @SuppressWarnings("unchecked")
                            Map<String, Object> sub = (Map<String, Object>) blockObj;
                            ctxForBlock.putAll(sub);
                        }

                        for (P blockP : buffer) {
                            String t = templateEngine.getParagraphText(blockP);
                            if (t == null) t = "";

                            // bỏ marker điều kiện
                            t = t.replace("{{?" + currentKey + "}}", "")
                                    .replace("{{/" + currentKey + "}}", "");
                            templateEngine.setParagraphText(blockP, t);

                            // 1) Nếu paragraph là placeholder ảnh duy nhất -> chèn ảnh
                            if (handleImagePlaceholder(pkg, blockP, ctxForBlock)) {
                                continue;
                            }

                            // 2) Còn lại xử lý scalar text như cũ
                            String textAfter = templateEngine.getParagraphText(blockP);
                            if (textAfter != null && textAfter.contains("{{")) {
                                textAfter = templateEngine.replaceScalars(textAfter, ctxForBlock);
                                if (textAfter.trim().isEmpty()) {
                                    templateEngine.deleteParagraph(pkg, blockP);
                                } else {
                                    templateEngine.setParagraphText(blockP, textAfter);
                                }
                            } else if (textAfter == null || textAfter.trim().isEmpty()) {
                                templateEngine.deleteParagraph(pkg, blockP);
                            }
                        }
                    } else {
                        // không thỏa điều kiện -> xoá toàn bộ block
                        for (P blockP : buffer) {
                            templateEngine.deleteParagraph(pkg, blockP);
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

            String replaced = templateEngine.replaceScalars(text, root);
            templateEngine.setParagraphText(p, replaced);
        }
    }
    private boolean handleImagePlaceholder(WordprocessingMLPackage pkg,
                                           P paragraph,
                                           Map<String, Object> ctx) throws Docx4JException {
        String txt = templateEngine.getParagraphText(paragraph);
        if (txt == null) return false;

        txt = txt.trim();
        Matcher m = IMAGE_ONLY_PLACEHOLDER.matcher(txt);
        if (!m.matches()) {
            return false; // paragraph không phải dạng "{{key}}" duy nhất
        }

        String key = m.group(1).trim(); // avatar, user.avatar, ...
        Object val = templateEngine.resolveKey(ctx, key);
        if (val == null) {
            // không có dữ liệu => xoá paragraph
            templateEngine.deleteParagraph(pkg, paragraph);
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
                templateEngine.deleteParagraph(pkg, paragraph);
            }
            return true;
        } catch (Exception e) {
            throw new RuntimeException("Error inserting image for key=" + key
                    + ", path=" + path, e);
        }
    }
    @SuppressWarnings("unchecked")
    private P createFlexImageParagraph(WordprocessingMLPackage pkg,
                                       Map<String, Object> block) throws Exception {
        Map<String, Object> imageData = (Map<String, Object>) block.get("imageData");
        if (imageData == null) return new P();

        String bucket = (String) imageData.get("bucket");
        String path = (String) imageData.get("path");

        byte[] bytes = templateEngine.loadImageBytes(bucket, path, "");
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

    // simple heuristic: path nhìn giống ảnh (URL, classpath, file ảnh)
    private boolean looksLikeImagePath(String path) {
        String p = path.toLowerCase(Locale.ROOT).trim();
        if (p.startsWith("http://") || p.startsWith("https://") || p.startsWith("classpath:")) {
            return true;
        }
        return p.endsWith(".png") || p.endsWith(".jpg") || p.endsWith(".jpeg")
                || p.endsWith(".gif") || p.endsWith(".bmp") || p.endsWith(".webp");
    }
}
