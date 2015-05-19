package com.powtechconsulting.mailmerge;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class WordMerger {

    public byte[] merge(String pathToFile, Map<String, String> mappings) {
        Context.getWmlObjectFactory();

        try {
            WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage.load(new File(pathToFile));


            MainDocumentPart mainDocumentPart = wordprocessingMLPackage.getMainDocumentPart();

            mainDocumentPart.variableReplace(new HashMap<String, String>(mappings));

            SaveToZipFile saver = new SaveToZipFile(wordprocessingMLPackage);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            saver.save(os);
            return os.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Problem merging file: " + pathToFile);
        }
    }
}
