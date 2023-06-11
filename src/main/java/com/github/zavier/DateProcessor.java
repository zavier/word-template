package com.github.zavier;

import org.apache.poi.xwpf.usermodel.*;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DateProcessor {

    public static void replaceDatePlaceholder(XWPFParagraph paragraph, Map<String, Object> map) {
        String text = paragraph.getText();

        // 判断是否包含日期占位符
        if (text.contains("${date:")) {
            for (XWPFRun run : paragraph.getRuns()) {
                String runText = run.getText(0);
                if (runText != null) {
                    // 使用正则表达式匹配日期占位符
                    Pattern pattern = Pattern.compile("\\$\\{date:(.*?):(.*?)\\}");
                    Matcher matcher = pattern.matcher(runText);

                    StringBuffer buffer = new StringBuffer();
                    while (matcher.find()) {
                        String placeholder = matcher.group(1);
                        String format = matcher.group(2);
                        String replacement = formatDate(format, (Date) map.get(placeholder));

                        // 替换日期占位符
                        matcher.appendReplacement(buffer, replacement);
                    }
                    matcher.appendTail(buffer);

                    // 更新段落中的文本
                    run.setText(buffer.toString(), 0);
                }
            }
        }
    }

    private static String formatDate(String format, Date date) {
        // 使用指定格式格式化当前日期
        SimpleDateFormat dateFormat = new SimpleDateFormat(format);
        return dateFormat.format(date);
    }
}
