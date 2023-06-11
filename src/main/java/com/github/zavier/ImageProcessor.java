package com.github.zavier;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ImageProcessor {
    public static void replaceImagePlaceholder(XWPFParagraph paragraph, Map<String, Object> imageMap) throws Exception {

        final List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            final XWPFRun run = runs.get(i);
            String text = run.getText(0);

            if (text != null) {
                String regex = "\\$\\{img:(.*?)\\}";
                if (text.matches(regex)) {
                    String placeholder = text.replaceAll(regex, "$1");

                    if (imageMap.containsKey(placeholder)) {
                        String imagePath = imageMap.get(placeholder).toString();
                        InputStream imageStream;

                        if (isRemoteImage(imagePath)) {
                            // 远程图片
                            imageStream = new URL(imagePath).openStream();
                        } else {
                            // 本地图片
                            imageStream = new FileInputStream(imagePath);
                        }

                        // 下载图片数据
                        byte[] imageData = IOUtils.toByteArray(imageStream);
                        imageStream.close();

                        // 插入图片
                        int pictureType = getImageType(imagePath);
                        int width = 100;
                        int height = 100;
                        run.setText("", 0);
                        run.addPicture(new ByteArrayInputStream(imageData), pictureType, "", Units.toEMU(width), Units.toEMU(height));

                        // 移除当前的 XWPFRun 对象
//                        paragraph.removeRun(i);
                    }
                }
            }
        }
    }


    private static boolean isRemoteImage(String imagePath) {
        return imagePath.startsWith("http://") || imagePath.startsWith("https://");
    }

    private static int getImageType(String imagePath) {
        if (imagePath.endsWith(".png")) {
            return XWPFDocument.PICTURE_TYPE_PNG;
        } else if (imagePath.endsWith(".jpeg") || imagePath.endsWith(".jpg")) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        } else if (imagePath.endsWith(".gif")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        } else {
            // 默认返回 PNG 类型，可根据需要进行调整
            return XWPFDocument.PICTURE_TYPE_PNG;
        }
    }
}
