import com.powtechconsulting.mailmerge.WordMerger;
import org.apache.commons.io.FileUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Test {
    public static void main(String[] args) throws Docx4JException, JAXBException, IOException {
        System.out.println("Starting");

        Map<String, String> mappings = new HashMap<String, String>();
        mappings.put("companyName", "Powtech Consulting Ltd");
        mappings.put("companyId", "1234");
        mappings.put("address", "32 Church Road\nGorslas\nLlanelli");
        byte[] mergeBytes = new WordMerger().merge("/Users/jamespowell//dev/jhipster/dryhome-crm/src/main/resources/merge-docs/dp_record.docx", mappings);
        FileUtils.writeByteArrayToFile(new File("out_" + System.currentTimeMillis() + ".docx"), mergeBytes);
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
