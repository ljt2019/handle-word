package com.tiger.itextpdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * PDF格式化输出工具类,【itextpdf-5.5.13.jar】
 * Create by tiger on 2019/6/8
 */
public class PdfUtil {

    private static final float borderWidth = 0.5f;//表格边框厚度
    private static final int totalRow = 33;//表格总行数
    private static final int interval = 5;//间距
    private static final int colCount = 3;//大列
    private static final int fixedHeight = 14;//固定表格高度

    public static void main(String[] args) throws IOException, DocumentException {
        test();
    }

    public static void test() throws IOException, DocumentException {
        String title = "成绩单";

        String outPath = "C:/AAAAA/outWord.doc";
        OutputStream os = new FileOutputStream(new File(outPath));

        Document doc = new Document();
        //设置纸张大小与方向
        doc.setPageSize(PageSize.A4.rotate());
        //设置边框
        doc.setMargins(10, 10, 10, 10);
        //输出
        PdfWriter writer = PdfWriter.getInstance(doc, os);
        doc.open();

//        RtfFont font = new RtfFont("宋体", 10.5f, Font.NORMAL);
//        Font font = new Font(setChinaFont(), 8, Font.NORMAL);// 设置字体

        // 定义字体 处理word和pdf,合并word,word转pdf等等
        FontFactoryImp ffi = new FontFactoryImp();
        // 注册全部默认字体目录，windows会自动找fonts文件夹的，返回值为注册到了多少字体
        ffi.registerDirectories();
        // 获取字体
        Font font = ffi.getFont("宋体", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, Font.UNDEFINED, null);

        // 增加一个段落,标题
        Paragraph paragraph = new Paragraph(title, font);
        paragraph.setAlignment(Element.ALIGN_CENTER);
        paragraph.setSpacingAfter(10);
        doc.add(paragraph);

        //创建表格1
        float[] widthArr1 = new float[]{204, 204, 204, 204};
        PdfPTable table1 = new PdfPTable(widthArr1.length);
        table1.setWidthPercentage(100);
        table1.setTotalWidth(819);// 设置表格的宽度,横向最宽度
        // 也可以每列分别设置宽度
        table1.setTotalWidth(widthArr1);

        table1.addCell(createCell("名称：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("院系：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("专业：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("学制：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("年级：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("班级：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("层次：", font, Element.ALIGN_CENTER,0));
        table1.addCell(createCell("学制1：", font, Element.ALIGN_CENTER,0));
        doc.add(table1);

        //创建表格2
        //每列分别设置宽度
        float[] widthArr = new float[]{123, 48, 34, 34, 34, 123, 48, 34, 34, 34, 123, 48, 34, 34, 34};
        PdfPTable table = new PdfPTable(widthArr.length);
        table.setWidthPercentage(100);
        table.setTotalWidth(819);// 设置表格的宽度,横向最宽度
        table.setTotalWidth(widthArr); //每列分别设置宽度
//        table.setLockedWidth(true);// 锁住宽度，注释点显示才正常，不知为何

        for (int i = 0; i < colCount; i++) {
            for (int j = 0; j < interval; j++) {
                //枚举中元素个数要等于 interval
                table.addCell(createCell(TableTitleEnum.getByCode(j).getMsg(), font, Element.ALIGN_CENTER));
            }
        }

        for (int i = 1; i <= colCount * interval * totalRow; i++) {
            table.addCell(createCell("", font, Element.ALIGN_CENTER));
        }

        //合并示例
        int count = colCount + 1;
        removeBorderWidth(table, 1, 0, count);
        setTableValueColSpan(table, 1, 0, "2016-2017学年第一学期", font);

        removeBorderWidth(table, 29, 10, count);
        setTableValueColSpan(table, 29, 10, "历年总学分", font);

        doc.add(table);//插入表格

        // 落款
        Paragraph para = new Paragraph("打印人：林先生       打印日期:" + new PdfDate(), font);
        para.setAlignment(Element.ALIGN_CENTER);
        para.setSpacingAfter(10);
        doc.add(para);
        Paragraph para1 = new Paragraph("注意：***********", font);
        para1.setAlignment(Element.ALIGN_LEFT);
        para1.setSpacingAfter(10);
        doc.add(para1);

        //新页
        doc.newPage();
        Paragraph paragraph1 = new Paragraph(title + "下一页", font);
        paragraph1.setAlignment(Element.ALIGN_CENTER);
        paragraph1.setSpacingAfter(10);
        doc.add(paragraph1);

        doc.close();
        writer.close();
        os.flush();
        os.close();
    }

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
                Phrase phrase = pdfPCell.getPhrase();
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
        PdfPCell cell = new PdfPCell();
        cell.setFixedHeight(fixedHeight);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setHorizontalAlignment(align);
        cell.setBorderWidth(borderWidth);
        cell.setPhrase(new Phrase(value, font));
        return cell;
    }

    public static PdfPCell createCell(String value, Font font, int align,float borderWidth) {
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
