package com.tiger.tiger.itextpdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPRow;
import com.itextpdf.text.pdf.PdfPTable;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * PDF格式化输出工具类,【itextpdf-5.5.13.jar】
 * Create by tiger on 2019/6/8
 */
public class PdfUtil {

    private PdfUtil() {
    }

    public static final float borderWidth = 0.5f;//表格边框厚度
    public static final int   totalRow    = 33;//表格总行数
    public static final int   interval    = 5;//间距
    public static final int   colCount    = 3;//大列
    public static final int   fixedHeight = 14;//固定表格高度

    /**
     * 隐藏表格边框
     *
     * @param table
     * @param rowIndex
     * @param cloIndex
     */
    public static void removeBorderWidth(PdfPTable table, int rowIndex, int cloIndex, int count) {
        //上侧
        for (int i = 1; i <= count; i++) {
            setTableBorderWidthTop(table, rowIndex, i + cloIndex);
        }

        //下侧
        for (int i = 1; i <= count; i++) {
            setTableBorderWidthBottom(table, rowIndex, i + cloIndex);
        }

        //左侧
        for (int i = 1; i <= count; i++) {
            setTableBorderWidthLeft(table, rowIndex, i + cloIndex);
        }

        //右侧
        for (int i = 1; i <= count; i++) {
            setTableBorderWidthRight(table, rowIndex, i + cloIndex);
        }
    }

    /**
     * 指定去除表格右侧线
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     */
    public static void setTableBorderWidthRight(PdfPTable table, int rowIndex, int colIndex) {
        ArrayList rows = table.getRows();
        PdfPRow row = (PdfPRow) rows.get(rowIndex);
        PdfPCell[] cells = row.getCells();
        PdfPCell cell = cells[colIndex];
        cell.setBorderWidthRight(0);
    }

    /**
     * 指定去除表格上侧线
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     */
    public static void setTableBorderWidthTop(PdfPTable table, int rowIndex, int colIndex) {
        ArrayList rows = table.getRows();
        PdfPRow row = (PdfPRow) rows.get(rowIndex);
        PdfPCell[] cells = row.getCells();
        PdfPCell cell = cells[colIndex];
        cell.setBorderWidthTop(0);
    }

    /**
     * 指定去除表格下侧线
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     */
    public static void setTableBorderWidthBottom(PdfPTable table, int rowIndex, int colIndex) {
        ArrayList rows = table.getRows();
        PdfPRow row = (PdfPRow) rows.get(rowIndex);
        PdfPCell[] cells = row.getCells();
        PdfPCell cell = cells[colIndex];
        cell.setBorderWidthBottom(0);
    }

    /**
     * 指定去除表格左侧线
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     */
    public static void setTableBorderWidthLeft(PdfPTable table, int rowIndex, int colIndex) {
        ArrayList rows = table.getRows();
        PdfPRow row = (PdfPRow) rows.get(rowIndex);
        PdfPCell[] cells = row.getCells();
        PdfPCell cell = cells[colIndex];
        cell.setBorderWidthLeft(0);
    }


    /**
     * 指定单元格 横向合并 赋值，合并 5 个单元格
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     * @param font
     */
    public static void setTableValueColSpan(PdfPTable table, int rowIndex, int colIndex,
                                            String value, Font font) {
        ArrayList rows = table.getRows();
        PdfPRow row = (PdfPRow) rows.get(rowIndex);
        PdfPCell[] cells = row.getCells();
        PdfPCell cell = cells[colIndex];
        Phrase newPhrase = Phrase.getInstance(2, value, font);
        cell.setColspan(5);
        cell.setBorderWidth(borderWidth);
        cell.setPhrase(newPhrase);
    }

    /**
     * 指定单元格赋值（定位到row与col）
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     * @param value
     * @param font
     */
    public static void setTableValue(PdfPTable table, int rowIndex, int colIndex,
                                     String value, Font font) {
        ArrayList rows = table.getRows();
        PdfPRow row = (PdfPRow) rows.get(rowIndex);
        PdfPCell[] cells = row.getCells();
        PdfPCell cell = cells[colIndex];
        Phrase newPhrase = Phrase.getInstance(2, value, font);
        cell.setBorderWidth(borderWidth);
        cell.setPhrase(newPhrase);
    }

    /**
     * 填充表格
     *
     * @param table
     * @param font
     */
    public static void setAllTableValue(PdfPTable table, Font font) {
        ArrayList rows = table.getRows();
        for (int i = 1; i < rows.size(); i++) {
            PdfPRow row = (PdfPRow) rows.get(i);
            PdfPCell[] cells = row.getCells();
            for (int j = 0; j < cells.length; j++) {
                PdfPCell pdfPCell = cells[j];
                Phrase newPhrase = Phrase.getInstance(2, "嘉佳" + i + j, font);
                pdfPCell.setPhrase(newPhrase);
            }
        }
    }

    /**
     * 生成一个表格并赋值
     *
     * @param value
     * @param font
     * @param align
     * @return
     */
    public static PdfPCell createCell(String value, Font font, int align) {
        return getPdfPCell(value, font, align, borderWidth);
    }

    public static PdfPCell createCell(String value, Font font, int align,float borderWidth) {
        return getPdfPCell(value, font, align, borderWidth);
    }

    public static PdfPCell getPdfPCell(String value, Font font, int align, float borderWidth) {
        PdfPCell cell = new PdfPCell();
        cell.setFixedHeight(fixedHeight);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setHorizontalAlignment(align);
        cell.setBorderWidth(borderWidth);
        cell.setPhrase(new Phrase(value, font));
        return cell;
    }

    /**
     * 汉字字体
     *
     * @return
     */
    public static BaseFont setChinaFont() {
        try {
            return BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 生成一个表格
     *
     * @param total    总列数
     * @param textFont 字体
     * @param data     表格数据     X行    Y列
     * @param doc      PDF文档对象
     * @throws DocumentException
     * @author hou_fx
     */
    public static void TableBule(int total, Font textFont, List<List<String>> data, Document doc) throws DocumentException {
        // 创建一个有N列的表格
        PdfPTable table = new PdfPTable(total);
        table.setPaddingTop(20);
        table.setSpacingAfter(20);
        table.setTotalWidth(530); //设置列宽
        // table.setTotalWidth(new float[]{ 100, 165, 100, 165 }); //设置列宽
        table.setLockedWidth(true); //锁定列宽
        PdfPCell cell;
        for (int i = 0; i < data.size(); i++) {  //遍历数据行   每个数据行都是一个list
            Iterator it = data.get(i).iterator();
            int count = 0;
            while (it.hasNext()) {               //遍历每行数据，每个数据都是一个单元格
                cell = new PdfPCell(new Phrase((String) it.next(), textFont));
                cell.setMinimumHeight(17); //设置单元格高度
                cell.setUseAscender(true); //设置可以居中
                //第一个单元格背景色
                if (count % 2 == 0) {
                    cell.setBackgroundColor(new BaseColor(231, 230, 230));
                }
                cell.setHorizontalAlignment(Element.ALIGN_LEFT); //左对齐
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE); //设置垂直居中
                table.addCell(cell);
                count++;
            }
        }
        doc.add(table);
    }

    /**
     * 生成一个表格
     *
     * @param total    总列数
     * @param textFont 字体
     * @param data     表格数据     X行    Y列
     * @param doc      PDF文档对象
     * @param colspan  第几列
     * @param rowspan  第几行
     * @param number   跨几列
     * @throws DocumentException
     * @author hou_fx
     */
    public static void TableColspan(int total, Font textFont, List<List<String>> data, Document doc, int[] rowspan, int[] colspan, int[] number) throws DocumentException {
        // 创建一个有N列的表格
        PdfPTable table = new PdfPTable(total);
        table.setPaddingTop(20);
        table.setSpacingAfter(20);
        table.setTotalWidth(530); //设置列宽
        // table.setTotalWidth(new float[]{ 100, 165, 100, 165 }); //设置列宽
        table.setLockedWidth(true); //锁定列宽
        PdfPCell cell;
        //数组下标
        int cos = 0;
        for (int i = 0; i < data.size(); i++) {  //遍历数据行   每个数据行都是一个list
            Iterator<String> it = data.get(i).iterator();
            int count = 0;
            while (it.hasNext()) {               //遍历每行数据，每个数据都是一个单元格
                cell = new PdfPCell(new Phrase(it.next(), textFont));
                cell.setMinimumHeight(17); //设置单元格高度
                cell.setUseAscender(true); //设置可以居中
                if (cos < rowspan.length && i == rowspan[cos] - 1 && count == colspan[cos] - 1) {
//					if (i==rowspan[cos]-1) {//行
//						if (count==colspan[cos]-1) {//列
                    cell.setColspan(number[cos]);//跨单元格
                    cos++;
//						}
//					}
                }
                cell.setHorizontalAlignment(Element.ALIGN_LEFT);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE); //设置垂直居中
                table.addCell(cell);
                count++;
            }
        }
        doc.add(table);
    }

}
