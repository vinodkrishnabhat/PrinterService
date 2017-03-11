package com.vkb.print;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;

import java.io.File;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.Map;

public class DocCreator {

    public static final int TWIPS_PER_INCH = 1440;

    public static void createDocFile(String fileName, Map<String, String> paragraphs) {
        try {
            File file = new File(fileName);
            FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());

            XWPFDocument doc = new XWPFDocument();

            for(String header: paragraphs.keySet()) {
                addHeader(doc, header);

                String fullContent = paragraphs.get(header);
                addContents(doc, fullContent);

                XWPFParagraph tempParagraph = doc.createParagraph();
                XWPFRun breakRun = tempParagraph.createRun();
            }

            doc.write(fos);
            fos.close();

            System.out.println(file.getAbsolutePath()+ " created successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void addContents(XWPFDocument doc, String fullContent) {
        String [] splitContent = fullContent.split("  +");
        XWPFTable table = doc.createTable();

        XWPFRun run;

        if(splitContent.length > 0) {
            XWPFTableRow tableRowOne = table.getRow(0);
            setHeight(tableRowOne);

            run = tableRowOne.getCell(0).getParagraphs().get(0).createRun();
            setRun(splitContent[0], run);

            if(splitContent.length > 1) {
                addAndSetCellValue(splitContent[1], tableRowOne);

                if(splitContent.length > 2) {
                    addAndSetCellValue(splitContent[2], tableRowOne);

                    if(splitContent.length > 3) {
                        addAndSetCellValue(splitContent[3], tableRowOne);
                    }
                }
            }
        }

        XWPFTableRow tableRow = null;

        for(int ctr = 4; ctr < splitContent.length; ctr++) {
            int col = ctr % 4;

            if(col == 0) {
                tableRow = table.createRow();
                setHeight(tableRow);
            }

            run = tableRow.getCell(col).getParagraphs().get(0).createRun();
            setRun(splitContent[ctr], run);
        }
    }

    private static void addAndSetCellValue(String value, XWPFTableRow tableRowOne) {
        XWPFTableCell cell;
        XWPFRun run;
        cell = tableRowOne.addNewTableCell();
        run = cell.getParagraphs().get(0).createRun();
        setRun(value, run);
    }

    private static void setHeight(XWPFTableRow tableRow) {
        tableRow.setHeight((int)(TWIPS_PER_INCH *2/10));
        tableRow.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT); //set w:hRule="exact"
    }

    private static void setRun(String value, XWPFRun run) {
        run.setFontSize(10);
        run.setText(value);
    }

    private static void addHeader(XWPFDocument doc, String header) {
        XWPFParagraph tempParagraph = doc.createParagraph();
        XWPFRun headerRun = tempParagraph.createRun();
        headerRun.setFontSize(12);
        headerRun.setBold(true);
        headerRun.setText(header);
    }

    public static void main(String[] args) {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("location1", "part one-2KGS     part two-2KGS");
        map.put("location2", "part one-2KGS     part two-2KGS     part three-2KGS     part four-2KGS");
        map.put("location3", "part one-2KGS     part two-2KGS     part three-2KGS     partfour-2KGS     part five-2KGS     part six-2KGS");
        map.put("location4", "part one-2KGS");
        map.put("location5", "part one-2KGS     part two-2KGS     part three-2KGS");
        map.put("location6", "");

        //create docx file
        DocCreator.createDocFile("C:/Temp/1.docx", map);

        //create doc file
        DocCreator.createDocFile("C:/Temp/2.doc", map);
    }
}
