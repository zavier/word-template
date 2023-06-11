package com.github.zavier;

import org.apache.poi.xwpf.usermodel.*;

import java.util.List;
import java.util.Map;

public class StringProcessor {

    public static void replaceParagraphPlaceholders(XWPFParagraph paragraph, Map<String, Object> placeholderMap) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String text = run.getText(0);
            if (text != null) {
                StringBuilder stringBuilder = new StringBuilder(text);
                for (Map.Entry<String, Object> entry : placeholderMap.entrySet()) {
                    String placeholder = "${string:" + entry.getKey() + "}";
                    int index = stringBuilder.indexOf(placeholder);
                    while (index != -1) {
                        stringBuilder.replace(index, index + placeholder.length(), entry.getValue().toString());
                        index = stringBuilder.indexOf(placeholder, index + entry.getValue().toString().length());
                    }
                }
                run.setText(stringBuilder.toString(), 0);
            }
        }
    }

}
