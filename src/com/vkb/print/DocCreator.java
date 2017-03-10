package com.vkb.print;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.Map;

public class DocCreator {
    public static void createDocFile(String fileName, Map<String, String> paragraphs) {
        try {
            File file = new File(fileName);
            FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());

            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph tempParagraph = doc.createParagraph();

            for(String header: paragraphs.keySet()) {
                XWPFRun headerRun = tempParagraph.createRun();
                headerRun.setFontSize(12);
                headerRun.setBold(true);
                headerRun.setText(header);
                headerRun.addBreak();

                XWPFRun contentRun = tempParagraph.createRun();
                contentRun.setFontSize(12);
                contentRun.setBold(false);
                contentRun.setText(paragraphs.get(header));
                contentRun.addBreak();
                contentRun.addBreak();
            }

            doc.write(fos);
            fos.close();

            System.out.println(file.getAbsolutePath()+ " created successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static void main(String[] args) {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("location1", "Str1");
        map.put("location2", "Str2");
        map.put("location3", "Str3");
        map.put("location4", "Str4");
        map.put("location5", "Str5");

        //create docx file
        DocCreator.createDocFile("C:/Temp/1.docx", map);

        //create doc file
        DocCreator.createDocFile("C:/Temp/2.doc", map);

    }
}
