package com.powtechconsulting.mailmerge;


import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

public class DocControlMerger {
    public byte[] merge(String pathToFile, String xmlFilename) throws Docx4JException, FileNotFoundException {
        // Load input_template.docx
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(pathToFile));

        // Open the xml stream
        FileInputStream xmlStream = new FileInputStream(new File(xmlFilename));

        // Do the binding:
        // FLAG_NONE means that all the steps of the binding will be done,
        // otherwise you could pass a combination of the following flags:
        // FLAG_BIND_INSERT_XML: inject the passed XML into the document
        // FLAG_BIND_BIND_XML: bind the document and the xml (including any OpenDope handling)
        // FLAG_BIND_REMOVE_SDT: remove the content controls from the document (only the content remains)
        // FLAG_BIND_REMOVE_XML: remove the custom xml parts from the document

        //Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_NONE);
        //If a document doesn't include the Opendope definitions, eg. the XPathPart,
        //then the only thing you can do is insert the xml
        //the example document binding-simple.docx doesn't have an XPathPart....
        Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_BIND_INSERT_XML & Docx4J.FLAG_BIND_BIND_XML);

        SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        saver.save(os);
        return os.toByteArray();
    }
}
