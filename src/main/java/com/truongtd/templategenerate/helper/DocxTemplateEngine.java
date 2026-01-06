package com.truongtd.templategenerate.helper;

import com.truongtd.templategenerate.request.GenerateTemplateRequest;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.wml.Body;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner;
import org.docx4j.wml.Tr;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.truongtd.templategenerate.util.StringUtils.DEFAULT_FONT;
import static com.truongtd.templategenerate.util.StringUtils.DEFAULT_FONT_SIZE;
import static com.truongtd.templategenerate.util.StringUtils.SCALAR;

@Slf4j
public class DocxTemplateEngine {

    private final DocxStyleHelper docxStyleHelper = new DocxStyleHelper(DEFAULT_FONT, DEFAULT_FONT_SIZE);

    public String replaceScalars(String text, Map<String, Object> ctx) {
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

    @SuppressWarnings("unchecked")
    public Object resolveKey(Map<String, Object> root, String key) {
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

    public String getParagraphText(P p) {
        StringBuilder sb = new StringBuilder();

        boolean inField = false;
        boolean inResult = false; // chỉ lấy text sau SEPARATE

        for (Object o : p.getContent()) {
            Object u = org.docx4j.XmlUtils.unwrap(o);
            if (!(u instanceof org.docx4j.wml.R)) continue;

            org.docx4j.wml.R r = (org.docx4j.wml.R) u;

            // detect fldChar in this run
            org.docx4j.wml.STFldCharType fldType = null;
            for (Object rc : r.getContent()) {
                Object ru = org.docx4j.XmlUtils.unwrap(rc);
                if (ru instanceof org.docx4j.wml.FldChar) {
                    fldType = ((org.docx4j.wml.FldChar) ru).getFldCharType();
                    break;
                }
            }

            if (fldType == org.docx4j.wml.STFldCharType.BEGIN) {
                inField = true;
                inResult = false;
                continue;
            }
            if (fldType == org.docx4j.wml.STFldCharType.SEPARATE) {
                inResult = true;
                continue;
            }
            if (fldType == org.docx4j.wml.STFldCharType.END) {
                inField = false;
                inResult = false;
                continue;
            }

            // collect visible text:
            // - ngoài field => lấy
            // - trong field => chỉ lấy khi đang ở result (sau SEPARATE)
            boolean canTake = !inField || inResult;

            if (!canTake) continue;

            for (Object rc : r.getContent()) {
                Object ru = org.docx4j.XmlUtils.unwrap(rc);
                if (ru instanceof org.docx4j.wml.Text) {
                    String v = ((org.docx4j.wml.Text) ru).getValue();
                    if (v != null) sb.append(v);
                }
            }
        }

        return sb.toString();
    }

    public void setParagraphText(P p, String newText) {
        docxStyleHelper.setParagraphTextStyled(p, newText);
    }

    public void deleteParagraph(WordprocessingMLPackage pkg, P p) throws Docx4JException {
        Body body = pkg.getMainDocumentPart().getContents().getBody();
        body.getContent().remove(p);
    }

    public boolean isTruthy(Object value) {
        if (value == null) return false;
        if (value instanceof Boolean) return (Boolean) value;
        if (value instanceof String ) {
            String s = (String) value;
            String x = s.trim();
            if (x.isEmpty()) return false;
            if ("false".equalsIgnoreCase(x)) return false;
            if ("true".equalsIgnoreCase(x)) return true;
            return true;
        }
        if (value instanceof Collection<?>) {
            Collection<?> c = (Collection<?>) value;
            return !c.isEmpty();
        }
        if (value instanceof Map<?, ?>) {
            Map<?,?> m = (Map<?, ?>) value;
            return !m.isEmpty();
        }
        return true;
    }

    public void applyVerticalMergeContinueAndClear(Tr row) {
        for (Object cellObj : row.getContent()) {
            Object cu = XmlUtils.unwrap(cellObj);
            if (!(cu instanceof Tc)) {
                continue;
            }
            Tc tc = (Tc) cu;

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

    public void clearNonListCellsForSubsequentRow(Tr row, String listKey) {
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

    public String getTcText(Tc tc) {
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

    public boolean isRowBlank(Tr row) {
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
    public boolean containsDrawing(Tc tc) {
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

    public byte[] downloadImageStrict(String url) throws IOException {
        HttpURLConnection conn = (HttpURLConnection) new URL(url).openConnection();
        conn.setConnectTimeout(10_000);
        conn.setReadTimeout(20_000);
        conn.setInstanceFollowRedirects(true);
        conn.setRequestProperty("User-Agent", "Mozilla/5.0");

        int code = conn.getResponseCode();
        if (code != 200) {
            throw new IOException("HTTP " + code + " when downloading image");
        }

        String ct = conn.getContentType();
        if (ct != null && !ct.toLowerCase(Locale.ROOT).startsWith("image/")) {
            throw new IOException("Not image content-type: " + ct);
        }

        try (InputStream is = conn.getInputStream()) {
            byte[] bytes = IOUtils.toByteArray(is);
            if (!looksLikeImage(bytes)) {
                throw new IOException("Downloaded content is not image (magic bytes mismatch)");
            }
            return bytes;
        }
    }

    public byte[] normalizeToPngStrict(byte[] input) throws IOException {
        if (input == null || input.length == 0)
            throw new IOException("Empty image bytes");

        BufferedImage img = ImageIO.read(new ByteArrayInputStream(input));
        if (img == null)
            throw new IOException("ImageIO cannot decode image");

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ImageIO.write(img, "png", bos);
        return bos.toByteArray();
    }

    private boolean looksLikeImage(byte[] b) {
        return looksLikePng(b) || looksLikeJpeg(b) || looksLikeGif(b) || looksLikeBmp(b);
    }
    private boolean looksLikePng(byte[] b) {
        return b != null && b.length > 8 &&
                b[0]==(byte)0x89 && b[1]==0x50 && b[2]==0x4E && b[3]==0x47;
    }
    private boolean looksLikeJpeg(byte[] b) {
        return b != null && b.length > 2 &&
                b[0]==(byte)0xFF && b[1]==(byte)0xD8;
    }
    private boolean looksLikeGif(byte[] b) {
        return b != null && b.length > 6 &&
                b[0]=='G' && b[1]=='I' && b[2]=='F';
    }
    private boolean looksLikeBmp(byte[] b) {
        return b != null && b.length > 2 &&
                b[0]=='B' && b[1]=='M';
    }

    public byte[] loadImageBytes(String bucket, String path, String imageBaseDir) {
        if (path == null || path.trim().isEmpty()) return new byte[0];
        path = path.trim();

        try {
            byte[] raw;

            // ========= 1) URL =========
            if (path.startsWith("http://") || path.startsWith("https://")) {
                raw = downloadImageStrict(path);   // check HTTP + content-type + magic bytes
                return normalizeToPngStrict(raw);  // chặn docx4j convertToPNG
            }

            // ========= 2) CLASSPATH =========
            if (path.startsWith("classpath:")) {
                String cp = path.substring("classpath:".length());
                String res = cp.startsWith("/") ? cp.substring(1) : cp;

                try (InputStream is = getClass().getClassLoader().getResourceAsStream(res)) {
                    if (is == null) return new byte[0];
                    raw = IOUtils.toByteArray(is);
                    return normalizeToPngStrict(raw);
                }
            }

            // ========= 3) LOCAL FILE =========
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

            if (!Files.exists(filePath)) return new byte[0];

            raw = Files.readAllBytes(filePath);
            return normalizeToPngStrict(raw);

        } catch (Exception e) {
            log.warn("Cannot load image. bucket={}, path={}", bucket, path, e);
            return new byte[0];
        }
    }
    public void fixLocationDateLayout(WordprocessingMLPackage pkg) {
        final String token = "{{applicationInfo.locationDateText}}";

        // duyệt tất cả paragraph trong main document
        new org.docx4j.TraversalUtil(pkg.getMainDocumentPart().getContent(),
                new org.docx4j.TraversalUtil.CallbackImpl() {
                    @Override
                    public java.util.List<Object> apply(Object o) {
                        Object u = org.docx4j.XmlUtils.unwrap(o);
                        if (!(u instanceof org.docx4j.wml.P)) return null;

                        org.docx4j.wml.P p = (org.docx4j.wml.P) u;

                        // NOTE: dùng hàm getParagraphTextVisibleOnly của anh (field-aware)
                        String txt = getParagraphText(p);
                        if (txt != null && txt.contains(token)) {
                            ensureSpacingBefore(p, 900); // 900 twips = 45pt (anh có thể chỉnh 600/720/900)
                        }
                        return null;
                    }
                }
        );
    }

    private void ensureSpacingBefore(org.docx4j.wml.P p, int beforeTwips) {
        org.docx4j.wml.ObjectFactory f = org.docx4j.jaxb.Context.getWmlObjectFactory();

        org.docx4j.wml.PPr ppr = p.getPPr();
        if (ppr == null) {
            ppr = f.createPPr();
            p.setPPr(ppr);
        }

        org.docx4j.wml.PPrBase.Spacing sp = ppr.getSpacing();
        if (sp == null) {
            sp = f.createPPrBaseSpacing();
            ppr.setSpacing(sp);
        }

        java.math.BigInteger current = sp.getBefore();
        java.math.BigInteger target = java.math.BigInteger.valueOf(beforeTwips);

        // chỉ tăng (không giảm) để tránh ảnh hưởng layout khác nếu template đã set sẵn
        if (current == null || current.compareTo(target) < 0) {
            sp.setBefore(target);
        }
    }

//    public void fixSoAndLocationDateNotOverlapLogo(WordprocessingMLPackage pkg) {
//
//        final Tbl[] targetTable = new Tbl[1];
//
//        new TraversalUtil(pkg.getMainDocumentPart().getContent(),
//                new TraversalUtil.CallbackImpl() {
//                    @Override
//                    public List<Object> apply(Object o) {
//                        if (targetTable[0] != null) return null;
//
//                        Object u = XmlUtils.unwrap(o);
//                        if (u instanceof Tbl) {
//                            Tbl tbl = (Tbl) u;
//                            String text = org.docx4j.TextUtils.getText(tbl);
//
//                            if (text != null &&
//                                    (text.contains("Số")
//                                            || text.contains("{{applicationInfo.locationDateText}}"))) {
//                                targetTable[0] = tbl;
//                            }
//                        }
//                        return null;
//                    }
//                }
//        );
//
//        //  BƯỚC 2: insert SAU traversal
//        if (targetTable[0] != null) {
//            insertSpacerParagraphBefore(targetTable[0], 900);
//        }
//    }
//
//    private void insertSpacerParagraphBefore(Tbl tbl, int beforeTwips) {
//        Object parent = tbl.getParent();
//        if (!(parent instanceof ContentAccessor)) return;
//
//        ContentAccessor ca = (ContentAccessor) parent;
//        List<Object> content = ca.getContent();
//
//        int idx = -1;
//        for (int i = 0; i < content.size(); i++) {
//            if (XmlUtils.unwrap(content.get(i)) == tbl) {
//                idx = i;
//                break;
//            }
//        }
//        if (idx < 0) return;
//
//        ObjectFactory f = Context.getWmlObjectFactory();
//        P spacer = f.createP();
//
//        PPr ppr = f.createPPr();
//        PPrBase.Spacing sp = f.createPPrBaseSpacing();
//        sp.setBefore(BigInteger.valueOf(beforeTwips));
//        ppr.setSpacing(sp);
//        spacer.setPPr(ppr);
//
//        content.add(idx, spacer);
//    }
}
