package com.word;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfDate;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.tiger.itextpdf.PdfUtil;
import com.tiger.itextpdf.TableTitleEnum;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 *
 * @author: tiger
 * @create: 2019-11-02 19:11
 */
public class ItextPdfTest {
    public static void main(String[] args) throws IOException, DocumentException {
        long start = System.currentTimeMillis();
        test();
        long end = System.currentTimeMillis();
        System.out.println("输出 Itext word成功！");
        System.out.println(end - start);
    }

    public static void test() throws IOException, DocumentException {
        String title = "成绩单";

        String outPath = "C:/AAAAA/outWord.pdf";
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

        table1.addCell(PdfUtil.createCell("名称：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("院系：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("专业：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("学制：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("年级：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("班级：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("层次：", font, Element.ALIGN_CENTER,0));
        table1.addCell(PdfUtil.createCell("学制1：", font, Element.ALIGN_CENTER,0));
        doc.add(table1);

        //创建表格2
        //每列分别设置宽度
        float[] widthArr = new float[]{123, 48, 34, 34, 34, 123, 48, 34, 34, 34, 123, 48, 34, 34, 34};
        PdfPTable table = new PdfPTable(widthArr.length);
        table.setWidthPercentage(100);
        table.setTotalWidth(819);// 设置表格的宽度,横向最宽度
        table.setTotalWidth(widthArr); //每列分别设置宽度
//        table.setLockedWidth(true);// 锁住宽度，注释点显示才正常，不知为何

        for (int i = 0; i < PdfUtil.colCount; i++) {
            for (int j = 0; j < PdfUtil.interval; j++) {
                //枚举中元素个数要等于 interval
                table.addCell(PdfUtil.createCell(TableTitleEnum.getByCode(j).getMsg(), font, Element.ALIGN_CENTER));
            }
        }

        for (int i = 1; i <= PdfUtil.colCount * PdfUtil.interval * PdfUtil.totalRow; i++) {
            table.addCell(PdfUtil.createCell("", font, Element.ALIGN_CENTER));
        }

        //合并示例
        int count = PdfUtil.colCount + 1;
        PdfUtil.removeBorderWidth(table, 1, 0, count);
        PdfUtil.setTableValueColSpan(table, 1, 0, "2016-2017学年第一学期", font);

        PdfUtil.removeBorderWidth(table, 29, 10, count);
        PdfUtil.setTableValueColSpan(table, 29, 10, "历年总学分", font);

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

}
