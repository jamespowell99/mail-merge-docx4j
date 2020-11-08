package com.powtechconsulting.mailmerge;

import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;

public class WordMergerDamprooferDetailsTest {
    private WordMergerDamprooferDetails wordMergerDamprooferDetails = new WordMergerDamprooferDetails();
    @Test
    public void testPdf() throws IOException {
        byte[] bytes = wordMergerDamprooferDetails.create(getMappings(), true);
        File outputFile = new File("dp-details_out_" + System.currentTimeMillis() + ".pdf");
        FileUtils.writeByteArrayToFile(outputFile, bytes);
        System.out.println("Wrote to " + outputFile);
    }

    @Test
    public void testWordDoc() throws IOException {
        byte[] bytes = wordMergerDamprooferDetails.create(getMappings(), false);
        File outputFile = new File("dp-details_out_" + System.currentTimeMillis() + ".docx");
        FileUtils.writeByteArrayToFile(outputFile, bytes);
        System.out.println("Wrote to " + outputFile);
    }

    private HashMap<String, String> getMappings() {
        HashMap<String, String> mappings = new HashMap<>();
        mappings.put("companyId", "3412");
        mappings.put("companyName", "N + B etc");
        mappings.put("address1", "line1");
        mappings.put("address2", "line2");
        mappings.put("address3", "line3");
        mappings.put("address4", "");
        mappings.put("address5", "");
        mappings.put("tel", "01269 841314");
        mappings.put("mob", "07800 806718");
        mappings.put("contact", "Mr James Powell");
        mappings.put("products", "275");
        mappings.put("price", "DRY");

        mappings.put("notes", "Some notes\nwhich\n\n\n\nwhich span\n across multiple \nlines");
        return mappings;
    }

}
