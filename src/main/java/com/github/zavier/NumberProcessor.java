package com.github.zavier;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.text.DecimalFormat;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class NumberProcessor {

    private static final String NUMBER_PLACEHOLDER_REGEX = "\\$\\{number:(\\w+)(?:,(\\d+)(\\+?))?}";
    static Pattern pattern = Pattern.compile(NUMBER_PLACEHOLDER_REGEX);

    public static void processParagraph(XWPFParagraph paragraph, Map<String, Object> data) {
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);

            if (text != null) {
                Matcher matcher = pattern.matcher(text);

                while (matcher.find()) {
                    String placeholder = matcher.group(0);
                    String fieldName = matcher.group(1);
                    String decimalPlacesStr = matcher.group(2);
                    String plusZeroFlag = matcher.group(3);

                    if (data.containsKey(fieldName)) {
                        double value = Double.parseDouble(data.get(fieldName).toString());
                        String formattedValue = formatNumber(value, decimalPlacesStr, plusZeroFlag);
                        text = text.replace(placeholder, formattedValue);
                    }
                }

                run.setText(text, 0);
            }
        }
    }

    private static String formatNumber(double value, String decimalPlacesStr, String plusZeroFlag) {
        int decimalPlaces = 2; // 默认保留两位小数

        if (decimalPlacesStr != null) {
            decimalPlaces = Integer.parseInt(decimalPlacesStr);
        }

        String pattern = "#";
        if (decimalPlaces > 0) {
            pattern += ".";
            for (int i = 0; i < decimalPlaces; i++) {
                pattern += "#";
            }
        }

        if (plusZeroFlag != null && plusZeroFlag.equals("+")) {
            pattern = pattern.replace("#", "0");
        }

        DecimalFormat decimalFormat = new DecimalFormat(pattern);
        return decimalFormat.format(value);
    }
}
