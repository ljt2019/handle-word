package com.tiger.aspose;

import com.aspose.words.License;
import com.aspose.words.SaveFormat;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * doc转pdf 【aspose-words-jdk16-14.6.0.0.jar】
 * Create by tiger on 2019/6/7
 */
public class Doc2PDF {

    public static void main(String[] args) {
        String sourcePath = "C:\\Users\\tiger\\Desktop\\研发三部-林继泰培训作业.docx";
        String targetPath = "C:\\Users\\tiger\\Desktop\\研发三部-林继泰培训作业.pdf";
        doc2pdf(sourcePath, targetPath);
    }

    /**
     * word文档转pdf
     *
     * @param wordPath word路径
     * @param pdfPath  pdf路径
     */
    public static void doc2pdf(String wordPath, String pdfPath) {
        // 验证License 若不验证则转化出的pdf文档会有水印产生

//        if (!getLicense()) {
//            return;
//        }
        FileOutputStream os = null;
        try {
            // 新建的PDF文件路径
            File file = new File(pdfPath);
            os = new FileOutputStream(file);
            // 要转换的word文档的路径
            com.aspose.words.Document doc =
                    new com.aspose.words.Document(wordPath);
            //设置字体
//            FontSettings fontSettings = new FontSettings();
//            fontSettings.setFontsFolder("/usr/share/fonts", true);
            // 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
            doc.save(os, SaveFormat.PDF);
            os.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 注册word2pdf转换工具
     *
     * @return
     */
    private static boolean getLicense() {
        boolean result = false;
        try {
            // 获取并配置license.xml
            String s = "<License>\n" +
                    "  <Data>\n" +
                    "    <Products>\n" +
                    "      <Product>Aspose.Total for Java</Product>\n" +
                    "      <Product>Aspose.Words for Java</Product>\n" +
                    "    </Products>\n" +
                    "    <EditionType>Enterprise</EditionType>\n" +
                    "    <SubscriptionExpiry>20991231</SubscriptionExpiry>\n" +
                    "    <LicenseExpiry>20991231</LicenseExpiry>\n" +
                    "    <SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber>\n" +
                    "  </Data>\n" +
                    "  <Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature>\n" +
                    "</License>";
            ByteArrayInputStream is = new ByteArrayInputStream(s.getBytes());
            License license = new License();
            license.setLicense(is);
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}

