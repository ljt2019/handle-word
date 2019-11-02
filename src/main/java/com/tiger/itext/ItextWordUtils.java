package com.tiger.itext;

import com.lowagie.text.*;

import com.lowagie.text.rtf.style.RtfFont;

import java.awt.*;

/**
 * 创建word文档 步骤:
 * 1,建立文档
 * 2,创建一个书写器
 * 3,打开文档
 * 4,向文档中写入数据
 * 5,关闭文档
 * add by linjitai on 20191028
 */
public class ItextWordUtils {

    private ItextWordUtils() {
    }

    /**
     * 获得一个表格(不指定总行数)
     *
     * @param width     表格占页面宽度比例
     * @param widths    每列宽度比例
     * @param padding   每行子格高度
     * @param alignment 居中模式
     * @return
     * @throws DocumentException
     */
    public static Table getTable(int width, int[] widths, float padding, int alignment) throws DocumentException {
        //创建表格对象
        Table table = new Table(widths.length);
        //设置表格占页面宽度比例
        table.setWidth(width);
        //设置每列宽度比例
        table.setWidths(widths);
        //设置每行子格高度
        table.setPadding(padding);
        //设置居中模式 Element.ALIGN_CENTER
        table.setAlignment(alignment);
        //设置表格线宽
//        table.setBorderWidth(0);
        //设置表格线宽颜色
        table.setBorderColor(Color.BLACK);
        //设置自动填满
//        table.setAutoFillEmptyCells(true);
        return table;
    }

    /**
     * 获得一个表格(指定总行数)
     *
     * @param width     表格占页面宽度比例
     * @param widths    每列宽度比例
     * @param padding   每行子格高度
     * @param alignment 居中模式
     * @param row       表格行数
     * @return
     * @throws DocumentException
     */
    public static Table getTable(int width, int row, float[] widths, float padding, int alignment) throws DocumentException {
        //创建表格对象
        Table table = new Table(widths.length, row);
        //设置表格占页面宽度比例
        table.setWidth(width);
        //设置每列宽度比例
        table.setWidths(widths);
        //设置每行子格高度
        table.setPadding(padding);
        //设置居中模式 Element.ALIGN_CENTER
        table.setAlignment(alignment);
        //设置表格线宽
        table.setBorderWidth(0.2f);
        //设置表格线宽颜色
        table.setBorderColor(Color.BLACK);
        //设置自动填满,不要设置这个
//        table.setAutoFillEmptyCells(true);
        return table;
    }

    /**
     * 向表格子指定位置（column,row）填充信息
     *
     * @param table     表格对象
     * @param column    列索引(从0开始)
     * @param row       行索引(从0开始)
     * @param content   内容
     * @param size      字体大小，例如 12
     * @param style     字体样式，例如 Font.BOLD
     * @param alignment 居中模式，例如 Element.ALIGN_LEFT
     */
    public static void fillCell(Table table, int column, int row, String content, float size, int style, int alignment, int borderWidthTop, int borderWidthBottom, int borderWidthLeft, int borderWidthRight) {
        try {
            Cell cell = getCell(borderWidthTop, borderWidthBottom, borderWidthLeft, borderWidthRight);
            cell.add(getParagraph(content, size, style, alignment));
            table.addCell(cell, column, row);
        } catch (BadElementException e) {
            e.printStackTrace();
        }
    }

    /**
     * 向表格子指定位置（column,row）填充信息
     *
     * @param table     表格对象
     * @param column    列索引(从0开始)
     * @param row       行索引(从0开始)
     * @param content   内容
     * @param size      字体大小，例如 12
     * @param style     字体样式，例如 Font.BOLD
     * @param alignment 居中模式，例如 Element.ALIGN_LEFT
     */
    public static void fillCell(Table table, int column, int row, String content, float size, int style, int alignment) {
        try {
            Cell cell = getCell();
            cell.add(getParagraph(content, size, style, alignment));
            table.addCell(cell, column, row);
        } catch (BadElementException e) {

//            e.printStackTrace();
        }
    }

    /**
     * 填充单元格
     * 右侧追加，到达列数则下一行追加
     *
     * @param table     表格对象
     * @param content   内容
     * @param size      字体大小，例如 12
     * @param style     字体样式，例如 Font.BOLD
     * @param alignment 居中模式，例如 Element.ALIGN_LEFT
     */
    public static void fillCell(Table table, String content, float size, int style, int alignment, int borderWidthTop, int borderWidthBottom, int borderWidthLeft, int borderWidthRight) {
        Cell cell = getCell(borderWidthTop, borderWidthBottom, borderWidthLeft, borderWidthRight);
        cell.add(getParagraph(content, size, style, alignment));
        table.addCell(cell);
    }

    /**
     * 填充单元格
     * 右侧追加，到达列数则下一行追加
     *
     * @param table     表格对象
     * @param content   内容
     * @param size      字体大小，例如 12
     * @param style     字体样式，例如 Font.BOLD
     * @param alignment 居中模式，例如 Element.ALIGN_LEFT
     */
    public static void fillCell(Table table, String content, float size, int style, int alignment) {
        Cell cell = getCell();
        cell.add(getParagraph(content, size, style, alignment));
        table.addCell(cell);
    }

    /**
     * 获得一个单元格子
     * 设置格子边框
     *
     * @return
     */
    public static Cell getCell(int borderWidthTop, int borderWidthBottom, int borderWidthLeft, int borderWidthRight) {
        Cell cell = new Cell();
        //水平居中
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        //垂直居中
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setBorderWidthTop(borderWidthTop);
        cell.setBorderWidthBottom(borderWidthBottom);
        cell.setBorderWidthLeft(borderWidthLeft);
        cell.setBorderWidthRight(borderWidthRight);

        //合并单元格
//        cell.setColspan(2);
//        cell.setRowspan(2);
        return cell;
    }

    /**
     * 获得一个单元格子
     *
     * @return
     */
    public static Cell getCell() {
        Cell cell = new Cell();
        //水平居中
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        //垂直居中
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        //合并单元格
//        cell.setColspan(2);
        return cell;
    }

    /**
     * 获得一个段落
     *
     * @param content
     * @param size      字体大小，例如 12
     * @param style     字体样式，例如 Font.BOLD
     * @param alignment 居中模式，例如 Element.ALIGN_LEFT
     * @return
     */
    public static Paragraph getParagraph(String content, float size, int style, int alignment) {
        Paragraph paragraph = new Paragraph(content, getFont(size, style));
        paragraph.setAlignment(alignment);
        return paragraph;
    }

    /**
     * 获取字体
     *
     * @param size  字体大小，例如 12
     * @param style 字体样式，例如 Font.BOLD
     * @return
     */
    public static RtfFont getFont(float size, int style) {
        //family = "仿宋"
        return new RtfFont("仿宋", size, style);
    }

}

