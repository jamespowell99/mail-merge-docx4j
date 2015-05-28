package com.powtechconsulting.mailmerge;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.utils.SingleTraversalUtilVisitorCallback;
import org.docx4j.wml.Body;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class WordMerger {
    //TODO user other methods of docx4j rather than variable replace. Much cleaner and much more suitable for
    public byte[] merge(String pathToFile, Map<String, String> mappings) {
        Context.getWmlObjectFactory();

        try {
            WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage.load(new File(pathToFile));

            VariablePrepare.prepare(wordprocessingMLPackage);

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
