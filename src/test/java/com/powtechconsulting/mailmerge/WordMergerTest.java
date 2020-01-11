package com.powtechconsulting.mailmerge;

import com.powtechconsulting.mailmerge.WordMerger;
import org.apache.commons.io.FileUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.junit.Ignore;
import org.junit.Test;

import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class WordMergerTest {
    @Test
    @Ignore
    public void test() throws Docx4JException, JAXBException, IOException {
        System.out.println("Starting");

        Map<String, String> mappings = new HashMap<String, String>();
//        mappings.put("currentDateTime", "12th May");
//        mappings.put("fullNameAddressLine1", "line1");
//        mappings.put("fullNameAddressLine2", "line2");
//        mappings.put("fullNameAddressLine3", "line3");
//        mappings.put("fullNameAddressLine4", "line4");
//        mappings.put("fullNameAddressLine5", "line5");
//        mappings.put("fullNameAddressLine6", "line6");
//        mappings.put("fullNameAddressLine7", "line7");
//        mappings.put("contact", "Mr James Powell");
//        mappings.put("invoiceNum", "abc123");
//        mappings.put("iNum", "abc123");
//        mappings.put("orderNumber", "xyz123");
//        mappings.put("james", "abc123");

        mappings.put("companyName", "myCompany");

        String fileName = this.getClass().getClassLoader().getResource("dampproofer-record.docx").getFile();
        byte[] mergeBytes = new WordMerger().merge(fileName, mappings);
        File outputFile = new File(fileName + "_out_" + System.currentTimeMillis() + ".docx");
        FileUtils.writeByteArrayToFile(outputFile, mergeBytes);
        System.out.println("Wrote to " + outputFile);
//        saveFile();
    }

    private static void saveFile() throws Docx4JException, JAXBException {
        Context.getWmlObjectFactory();

        WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage.load(new File("/Users/jamespowell/Downloads/testing.docx"));

        MainDocumentPart mainDocumentPart = wordprocessingMLPackage.getMainDocumentPart();

        HashMap<String, String> mappings = new HashMap<String, String>();
        mappings.put("colour", "green");
        mappings.put("icecream", "chocolate");

        mainDocumentPart.variableReplace(mappings);

        SaveToZipFile saver = new SaveToZipFile(wordprocessingMLPackage);
        saver.save(new File("out.docx"));
    }

}
