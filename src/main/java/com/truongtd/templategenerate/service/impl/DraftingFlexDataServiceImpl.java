package com.truongtd.templategenerate.service.impl;

import com.truongtd.templategenerate.helper.DocxStyleHelper;
import com.truongtd.templategenerate.helper.DocxTemplateEngine;
import com.truongtd.templategenerate.service.DraftingFlexDataService;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.truongtd.templategenerate.util.StringUtils.DEFAULT_FONT;
import static com.truongtd.templategenerate.util.StringUtils.DEFAULT_FONT_SIZE;

@Service
@Slf4j
public class DraftingFlexDataServiceImpl implements DraftingFlexDataService {
    @Value("${template.image.base-dir:}")   // optional base dir
    private String imageBaseDir;
    DocxTemplateEngine templateEngine = new DocxTemplateEngine();

    private final DocxStyleHelper docxStyleHelper = new DocxStyleHelper(DEFAULT_FONT, DEFAULT_FONT_SIZE);
    @Override
    public void processFlexData(WordprocessingMLPackage pkg,
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
                String txt = templateEngine.getParagraphText(p);
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
        docxStyleHelper.createParagraph(data);
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
                docxStyleHelper.createCell(text);
                P p = new P();
                R row = new R();
                Text valueText = new Text();
                valueText.setValue(text != null ? text : "");
                row.getContent().add(docxStyleHelper.createCell(valueText.getValue()));
                if (isBold) {
                    RPr rPr = new RPr();
                    rPr.setB(new BooleanDefaultTrue());
                    row.setRPr(rPr);
                }
                p.getContent().add(row);
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

        byte[] bytes = templateEngine.loadImageBytes(bucket, path, imageBaseDir);
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
}
