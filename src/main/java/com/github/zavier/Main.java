package com.github.zavier;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.*;

public class Main {


    public static void main(String[] args) throws Exception {

        final InputStream templateStream = Main.class.getClassLoader().getResourceAsStream("template.docx");

        // 创建模板数据
        Map<String, Object> placeholderMap = new HashMap<>();
        placeholderMap.put("name", "张三");
        placeholderMap.put("profession", "里斯");
        placeholderMap.put("ssname", "John Doe");
        placeholderMap.put("price", String.valueOf(123.45));
        placeholderMap.put("date", new Date());
        placeholderMap.put("image1", "https://zhengw-tech.com/images/netty-server.png");
        placeholderMap.put("image2", "https://zhengw-tech.com/images/jvm-class.png");

        try {
            XWPFDocument document = new XWPFDocument(templateStream);


            for (XWPFParagraph paragraph : document.getParagraphs()) {
                StringProcessor.replaceParagraphPlaceholders(paragraph, placeholderMap);
                NumberProcessor.processParagraph(paragraph, placeholderMap);
                DateProcessor.replaceDatePlaceholder(paragraph, placeholderMap);
                ImageProcessor.replaceImagePlaceholder(paragraph, placeholderMap);
            }

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            StringProcessor.replaceParagraphPlaceholders(paragraph, placeholderMap);
                            NumberProcessor.processParagraph(paragraph, placeholderMap);
                            DateProcessor.replaceDatePlaceholder(paragraph, placeholderMap);
                            ImageProcessor.replaceImagePlaceholder(paragraph, placeholderMap);
                        }
                    }
                }
            }

            FileOutputStream outputStream = new FileOutputStream("output.docx");
            document.write(outputStream);
            outputStream.close();

            System.out.println("Word template replaced successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}