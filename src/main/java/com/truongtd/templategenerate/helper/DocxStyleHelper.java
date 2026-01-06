package com.truongtd.templategenerate.helper;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;

import java.math.BigInteger;

public class DocxStyleHelper {
    private final ObjectFactory factory = Context.getWmlObjectFactory();
    private final String fontName;
    private final int fontSizeHalfPoints; // docx size = half-points (13pt => 26)

    public DocxStyleHelper(String fontName, int fontSizePt) {
        this.fontName = fontName;
        this.fontSizeHalfPoints = fontSizePt * 2;
    }

    /** Run properties: font + size */
    public RPr createRPr() {
        RPr rPr = factory.createRPr();

        RFonts rFonts = factory.createRFonts();
        rFonts.setAscii(fontName);
        rFonts.setHAnsi(fontName);
        rFonts.setCs(fontName);
        rPr.setRFonts(rFonts);

        HpsMeasure sz = new HpsMeasure();
        sz.setVal(BigInteger.valueOf(fontSizeHalfPoints));
        rPr.setSz(sz);
        rPr.setSzCs(sz);

        return rPr;
    }

    /** Create a run with styled text */
    public R createTextRun(String text) {
        R r = factory.createR();
        r.setRPr(createRPr());

        Text t = factory.createText();
        t.setValue(text == null ? "" : text);
        // Preserve spaces if needed
        t.setSpace("preserve");

        r.getContent().add(t);
        return r;
    }

    /** Ensure paragraph has 1 run with styled text (simple but stable for template-engine) */
    public void setParagraphTextStyled(P p, String text) {
        p.getContent().clear();
        p.getContent().add(createTextRun(text));
    }

    /** Create a paragraph with styled text */
    public P createParagraph(String text) {
        P p = factory.createP();
        p.getContent().add(createTextRun(text));
        return p;
    }

    /** Create a table cell with styled text */
    public Tc createCell(String text) {
        Tc tc = factory.createTc();
        P p = createParagraph(text);
        tc.getContent().add(p);
        return tc;
    }
}
