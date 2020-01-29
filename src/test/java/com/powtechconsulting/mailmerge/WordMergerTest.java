package com.powtechconsulting.mailmerge;

import com.powtechconsulting.mailmerge.WordMerger;
import com.sun.rmi.rmid.ExecPermission;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
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
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

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

        mappings.put("companyName", "myCompany22");
        mappings.put("companyId", "theCompanyId");
        mappings.put("notes", "line1\nline2\nline3");

        String fileName = this.getClass().getClassLoader().getResource("dampproofer-record.docx").getFile();
        byte[] mergeBytes = new WordMerger().merge(fileName, mappings);
        File outputFile = new File(fileName + "_out_" + System.currentTimeMillis() + ".docx");
        FileUtils.writeByteArrayToFile(outputFile, mergeBytes);
        System.out.println("Wrote to " + outputFile);
//        saveFile();
    }

    @Test
    public void testInvoice() throws Exception{
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");

        Map<String, String> map = new HashMap<>();
        map.put("invceNo", "abc123");
        map.put("orderNo", "sm12345");
        map.put("invoiceDate", LocalDate.now().format(formatter));
        map.put("ref", "Mr Dave Sharp");
        map.put("customerName", "Aquacure");
        List<String> address = Stream.of("3 Vicarage Lane", "Frampton On Severn", "GLOUCESTER", "GL2 7EE")
                .filter(x -> !StringUtils.isEmpty(x))
                .collect(Collectors.toList());
        for (int i = 0; i < 5; i++) {
            try {
                map.put("address" + (i+1), address.get(i));
            } catch (IndexOutOfBoundsException e) {
                map.put("address" + (i+1), "");
            }
        }

        map.put("delAddress1", "As Opposite");
        map.put("delAddress2", "");
        map.put("delAddress3", "");
        map.put("delAddress4", "");

        map.put("notes1", "");
        map.put("notes2", "");

        map.put("telNo", "01452 741277");



        for (int i = 1; i <= 5; i++) {
                map.put("item" + i, "test" + i);
                map.put("qt" + i, String.valueOf(i));
                map.put("price" + i, "1.99");
                map.put("total" + i, "2.99");
        }

        map.put("subtl", "8.99");
        map.put("vat", "2.99");
        map.put("total", "11.99");
        map.put("vrt", new DecimalFormat("#0.#").format(new BigDecimal(20)));


            map.put("paymentDetailsHeader", "Payment Details");
            map.put("paymentStatus", "Payment Received");

            map.put("paymentTypeLabel", "Type:");
            map.put("paymentType", "VISA");
            map.put("paymentDateLabel", "Date:");
            map.put("paymentDate", LocalDate.now().format(formatter));
            map.put("paymentAmountLabel", "Amount");
            String invoicePaymentAmount;
                invoicePaymentAmount = "£" + BigDecimal.TEN.setScale(2, BigDecimal.ROUND_UP).toString();
            map.put("paymentAmount", invoicePaymentAmount);
//            map.put("paymentDetailsHeader", "");
//            map.put("paymentStatus", "");
//
//            map.put("paymentTypeLabel", "");
//            map.put("paymentType", "");
//            map.put("paymentDateLabel", "");
//            map.put("paymentDate", "");
//            map.put("paymentAmountLabel", "");
//            map.put("paymentAmount", "");


        String fileName = this.getClass().getClassLoader().getResource("customer-invoice.docx").getFile();
        byte[] mergeBytes = new WordMerger().merge(fileName, map);
        File outputFile = new File(fileName + "_out_" + System.currentTimeMillis() + ".docx");
        FileUtils.writeByteArrayToFile(outputFile, mergeBytes);
        System.out.println("Wrote to " + outputFile);
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
