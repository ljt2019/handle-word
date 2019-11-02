package com.tiger.xwpf;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.List;
import java.util.Map;

/**
 * @description:
 * @author: tiger
 * @create: 2019-11-02 19:08
 */
public class WordTagTemplate {


    /**
     * 替换模板中的标签为实际的内容
     *
     * @param map
     */
    public static void replaceTag(XWPFDocument document,
                                  Map<String, String> map) {
        replaceParagraphs(document, map);
        replaceTables(document, map);
    }

    /**
     * 替换文本中的标签
     *
     * @param map key(待替换标签)-value(文本内容)
     */
    public static void replaceParagraphs(XWPFDocument document,
                                         Map<String, String> map) {
        List<XWPFParagraph> allXWPFParagraphs = document.getParagraphs();
        for (XWPFParagraph XwpfParagrapg : allXWPFParagraphs) {
            XWPFParagraphHandler XwpfParagrapgUtils =
                    new XWPFParagraphHandler(XwpfParagrapg);
            XwpfParagrapgUtils.replaceAll(map);
        }
    }

    /**
     * 替换表格中的标签
     *
     * @param map key(待替换标签)-value(文本内容)
     */
    public static void replaceTables(XWPFDocument document,
                                     Map<String, String> map) {
        List<XWPFTable> xwpfTables = document.getTables();
        for (XWPFTable xwpfTable : xwpfTables) {
            XWPFTableHandler xwpfTableUtils = new XWPFTableHandler(xwpfTable);
            xwpfTableUtils.replace(map);
        }
    }
}
