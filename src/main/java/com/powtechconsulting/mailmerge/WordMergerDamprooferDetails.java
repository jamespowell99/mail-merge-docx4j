package com.powtechconsulting.mailmerge;

import org.apache.commons.io.FileUtils;
import org.docx4j.Docx4J;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Text;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class WordMergerDamprooferDetails {
    public byte[] create(Map<String, String> mappings, boolean createPdf) {
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
            P pCompanyNameHeader = getCompanyNameHeader(factory, mappings.get("companyName"));

            wordMLPackage.getMainDocumentPart().addObject(pCompanyNameHeader);
            wordMLPackage.getMainDocumentPart().addObject(getFieldValue(factory, "Company ID Number: ", mappings.get("companyId")));
            wordMLPackage.getMainDocumentPart().addObject(getFieldValue(factory, "Company Name: ", mappings.get("companyName")));
            wordMLPackage.getMainDocumentPart().addObject(getBreak(factory));

            wordMLPackage.getMainDocumentPart().addObject(createAddressLinesFromList(factory, mappings));

            wordMLPackage.getMainDocumentPart().addObject(getFieldValue(factory, "Telephone: ", mappings.get("tel")));
            wordMLPackage.getMainDocumentPart().addObject(getFieldValue(factory, "Mobile: ", mappings.get("mobile")));
            wordMLPackage.getMainDocumentPart().addObject(getFieldValue(factory, "Contact: ", mappings.get("contact")));
            wordMLPackage.getMainDocumentPart().addObject(getBreak(factory));
            wordMLPackage.getMainDocumentPart().addObject(createPaidAndProducts(factory, mappings.get("paid"), mappings.get("products")));
            wordMLPackage.getMainDocumentPart().addObject(getBreak(factory));


            wordMLPackage.getMainDocumentPart().addObject(createNotes(factory, mappings.get("notes")));

                ByteArrayOutputStream os = new ByteArrayOutputStream();
            if (createPdf) {
                Docx4J.toPDF(wordMLPackage, os);
            } else {
                wordMLPackage.save(os);
            }
                return os.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("Problem creating file", e);
        }
    }

    //todo clean all this up
    private P getBreak(ObjectFactory factory) {
        P p = factory.createP();
        PPr pPr = factory.createPPr();
        ParaRPr paraRPr = factory.createParaRPr();
        pPr.setRPr(paraRPr);

        R r = factory.createR();
        RPr rPr = factory.createRPr();
        BooleanDefaultTrue booleanDefaultTrue = factory.createBooleanDefaultTrue();
        booleanDefaultTrue.setVal(false);
        rPr.setRtl(booleanDefaultTrue);
        r.setRPr(rPr);
        p.getContent().add(r);

        return p;
    }

    private P getCompanyNameHeader(ObjectFactory factory, String companyName) {
        P  p = factory.createP();
        PPr pPr = factory.createPPr();
        Jc jc = factory.createJc();
        jc.setVal(JcEnumeration.CENTER);
        pPr.setJc(jc);
        ParaRPr paraRPr = factory.createParaRPr();
        HpsMeasure hpsMeasure = factory.createHpsMeasure();
        hpsMeasure.setVal(BigInteger.valueOf(48));
        paraRPr.setSz(hpsMeasure);
        paraRPr.setSzCs(hpsMeasure);
        pPr.setRPr(paraRPr);
        p.setPPr(pPr);
        R r = factory.createR();
        RPr rPr = factory.createRPr();
        rPr.setSz(hpsMeasure);
        rPr.setSzCs(hpsMeasure);
        BooleanDefaultTrue booleanDefaultTrue = factory.createBooleanDefaultTrue();
        booleanDefaultTrue.setVal(false);
        rPr.setRtl(booleanDefaultTrue);
        r.setRPr(rPr);
        Text text = factory.createText();
//        JAXBElement<Text> textWrapped = factory.createRInstrText(text);
        r.getContent().add( text);
        text.setValue(companyName);
        text.setSpace( "preserve");
        p.getContent().add(r);
        return p;
    }

    private P getFieldValue(ObjectFactory factory, String fieldName, String fieldValue) {
        P  p = factory.createP();
        PPr pPr = factory.createPPr();
        p.setPPr(pPr);

        R r = factory.createR();
        RPr rPr = factory.createRPr();
        rPr.setB(factory.createBooleanDefaultTrue());
        rPr.setI(factory.createBooleanDefaultTrue());
        BooleanDefaultTrue booleanDefaultTrue = factory.createBooleanDefaultTrue();
        booleanDefaultTrue.setVal(false);
        rPr.setRtl(booleanDefaultTrue);
        r.setRPr(rPr);
        Text text = factory.createText();
//        JAXBElement<Text> textWrapped = factory.createRInstrText(text);
        r.getContent().add( text);
        text.setValue(fieldName);
        text.setSpace( "preserve");
        p.getContent().add(r);

        R r2 = factory.createR();
        RPr rPr2 = factory.createRPr();
        BooleanDefaultTrue booleanDefaultTrue2 = factory.createBooleanDefaultTrue();
        booleanDefaultTrue2.setVal(false);
        rPr2.setRtl(booleanDefaultTrue2);
        r2.setRPr(rPr2);
        Text text2 = factory.createText();
//        JAXBElement<Text> textWrapped2 = factory.createRInstrText(text2);
        r2.getContent().add( text2);
        text2.setValue(fieldValue);
        text2.setSpace( "preserve");

//        Br br = factory.createBr(); // this Br element is used break the current and go for next line
//        r2.getContent().add(br);

        p.getContent().add(r2);


        return p;
    }


    private P createAddressLinesFromList(ObjectFactory factory, Map<String, String> mappings) {
        P  p = factory.createP();
        PPr pPr = factory.createPPr();
        p.setPPr(pPr);

        R r = factory.createR();
        RPr rPr = factory.createRPr();
        rPr.setB(factory.createBooleanDefaultTrue());
        rPr.setI(factory.createBooleanDefaultTrue());
        BooleanDefaultTrue booleanDefaultTrue = factory.createBooleanDefaultTrue();
        booleanDefaultTrue.setVal(false);
        rPr.setRtl(booleanDefaultTrue);
        r.setRPr(rPr);
        Text text = factory.createText();
        r.getContent().add( text);
        text.setValue("Address:");
        text.setSpace( "preserve");
        Br br1 = factory.createBr(); // this Br element is used break the current and go for next line
        r.getContent().add(br1);
        p.getContent().add(r);

        R r2 = factory.createR();
        RPr rPr2 = factory.createRPr();
        BooleanDefaultTrue booleanDefaultTrue2 = factory.createBooleanDefaultTrue();
        booleanDefaultTrue2.setVal(false);
        rPr2.setRtl(booleanDefaultTrue2);
        r2.setRPr(rPr2);
        for (int i = 0; i < 5; i++) {
            Text text2 = factory.createText();
            r2.getContent().add( text2);
            text2.setValue(mappings.get("address" + (i+1)));
            text2.setSpace( "preserve");

            Br br = factory.createBr(); // this Br element is used break the current and go for next line
            r2.getContent().add(br);
        }


        p.getContent().add(r2);



        return p;
    }

    private P createNotes(ObjectFactory factory, String notes) {
        P  p = factory.createP();
        PPr pPr = factory.createPPr();
        p.setPPr(pPr);

        R r = factory.createR();
        RPr rPr = factory.createRPr();
        rPr.setB(factory.createBooleanDefaultTrue());
        rPr.setI(factory.createBooleanDefaultTrue());
        BooleanDefaultTrue booleanDefaultTrue = factory.createBooleanDefaultTrue();
        booleanDefaultTrue.setVal(false);
        rPr.setRtl(booleanDefaultTrue);
        r.setRPr(rPr);
        Text text = factory.createText();
        r.getContent().add( text);
        text.setValue("Notes:");
        text.setSpace( "preserve");
        Br br1 = factory.createBr(); // this Br element is used break the current and go for next line
        r.getContent().add(br1);
        p.getContent().add(r);

        R r2 = factory.createR();
        RPr rPr2 = factory.createRPr();
        BooleanDefaultTrue booleanDefaultTrue2 = factory.createBooleanDefaultTrue();
        booleanDefaultTrue2.setVal(false);
        rPr2.setRtl(booleanDefaultTrue2);
        r2.setRPr(rPr2);
        List<String> lines = Arrays.asList(notes.split("\n"));
        lines.forEach(l -> {
            Text text2 = factory.createText();
            r2.getContent().add( text2);
            text2.setValue(l);
            text2.setSpace( "preserve");

            Br br = factory.createBr(); // this Br element is used break the current and go for next line
            r2.getContent().add(br);
        });

        p.getContent().add(r2);



        return p;
    }

    private P createPaidAndProducts(ObjectFactory factory, String paid, String products) {
        P  p = factory.createP();
        PPr pPr = factory.createPPr();
        p.setPPr(pPr);

        R r = factory.createR();
        RPr rPr = factory.createRPr();
        rPr.setB(factory.createBooleanDefaultTrue());
        rPr.setI(factory.createBooleanDefaultTrue());
        BooleanDefaultTrue booleanDefaultTrue = factory.createBooleanDefaultTrue();
        booleanDefaultTrue.setVal(false);
        rPr.setRtl(booleanDefaultTrue);
        r.setRPr(rPr);
        Text text = factory.createText();
        r.getContent().add( text);
        text.setValue("Paid:");
        text.setSpace( "preserve");
        p.getContent().add(r);

        R r2 = factory.createR();
        RPr rPr2 = factory.createRPr();
        BooleanDefaultTrue booleanDefaultTrue2 = factory.createBooleanDefaultTrue();
        booleanDefaultTrue2.setVal(false);
        rPr2.setRtl(booleanDefaultTrue2);
        r2.setRPr(rPr2);
        Text text2 = factory.createText();
        r2.getContent().add( text2);
        text2.setValue(paid);
        text2.setSpace( "preserve");

        Br br = factory.createBr(); // this Br element is used break the current and go for next line
        r2.getContent().add(br);
        p.getContent().add(r2);


        R r3 = factory.createR();
        RPr rPr3 = factory.createRPr();
        rPr3.setB(factory.createBooleanDefaultTrue());
        rPr3.setI(factory.createBooleanDefaultTrue());
        BooleanDefaultTrue booleanDefaultTrue3 = factory.createBooleanDefaultTrue();
        booleanDefaultTrue3.setVal(false);
        rPr3.setRtl(booleanDefaultTrue3);
        r3.setRPr(rPr3);
        Text text3 = factory.createText();
        r3.getContent().add( text3);
        text3.setValue("Products:");
        text3.setSpace( "preserve");
        p.getContent().add(r3);

        R r4 = factory.createR();
        RPr rPr4 = factory.createRPr();
        BooleanDefaultTrue booleanDefaultTrue4 = factory.createBooleanDefaultTrue();
        booleanDefaultTrue4.setVal(false);
        rPr4.setRtl(booleanDefaultTrue2);
        r4.setRPr(rPr4);
        Text text4 = factory.createText();
        r4.getContent().add( text4);
        text4.setValue(products);
        text4.setSpace( "preserve");
        p.getContent().add(r4);




        return p;
    }


}
