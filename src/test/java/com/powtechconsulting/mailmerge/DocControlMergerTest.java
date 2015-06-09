package com.powtechconsulting.mailmerge;

import org.apache.commons.io.FileUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class DocControlMergerTest {

    @Ignore
    @Test
    public void test() throws IOException, Docx4JException {
        String fileName = this.getClass().getClassLoader().getResource("remcon_prod_lit.docx").getFile();
        String xmlFilename = this.getClass().getClassLoader().getResource("customer.xml").getFile();
        byte[] bytes = new DocControlMerger().merge(fileName, xmlFilename);

        File outputFile = new File(fileName + "_out_" + System.currentTimeMillis() + ".docx");
        FileUtils.writeByteArrayToFile(outputFile, bytes);
        System.out.println("Wrote to " + outputFile);
    }
}