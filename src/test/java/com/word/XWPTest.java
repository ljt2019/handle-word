package com.word;


import com.tiger.tiger.constant.Constants;
import com.tiger.tiger.enums.TitleEnum;
import com.tiger.tiger.enums.XWPFWidthEnum;
import com.tiger.tiger.xwpf.XWPFWordUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import java.io.*;
import java.util.List;

public class XWPTest {
    public XWPTest() {
        super();
    }

    public static void main(String[] args) throws IOException {
        long start = System.currentTimeMillis();
        String filePath = "c:/AAAAA/xwpf"+System.currentTimeMillis()+".docx";
        test(filePath);
        long end = System.currentTimeMillis();
        System.out.println("输出 XWPFDocument word成功！");
        System.out.println(end - start);
    }

    private static void testMerge() throws IOException {
        //默认：宋体（wps）/等线（office2016） 5号 两端对齐 单倍间距

        String template = "F:\\worldTest\\xjkxx_template.docx"; //模板路径
        String outPath = "F:\\worldTest\\合并.docx"; //写出路径

        File file = new File(template);
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis); //文档输入流

        List<XWPFTable> tableList = doc.getTables();

        XWPFTable table = tableList.get(1);
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;

        XWPFTableRow row = table.getRow(3); //获取某行
        XWPFTableCell cell = row.getCell(1); //获取行中某个表格

        paras = cell.getParagraphs(); //获取行中某个表格中所有段落

        //替换文字
        XWPFParagraph para = paras.get(0); //获取某段落
        List<XWPFRun> runs = para.getRuns();

        System.out.println("===runs.size()====" + runs.size());
        if (runs.size() <= 0) {
            System.out.println("===新建文字=====");
            XWPFRun run = para.insertNewRun(0);
            run.setFontFamily("宋体正文");
            run.setFontSize(7);
            run.setText("新建文字");
        } else {
            XWPFRun run = runs.get(0);
            run.setText("替字", 0);
            System.out.println("===替换=====" + run.toString());
        }

        //合并单元格
        XWPFWordUtils.mergeCellsHorizontal(table, 3, 2, 3); //

        //        XWPFParagraph para = cell.addParagraph();
        //        para.getParagraphText();
        ////       XWPFRun tempRun = p.getRuns().get(0);
        ////        tempRun.setText("", 0);
        //        XWPFRun run = para.createRun();
        //        run.setFontFamily("宋体");
        //        run.setFontSize(6);
        //        run.setText("字段说明");

        //        row.getCell(2).setText("字段说明");


        File fileOut = new File(outPath);
        FileOutputStream fos = new FileOutputStream(fileOut);
        BufferedOutputStream bos = new BufferedOutputStream(fos);
        doc.write(bos);
        //关闭流
        fos.flush();
        bos.flush();
        bos.close();
        fos.close();
        fis.close();

        System.out.println("====== 导出成功 ======");
    }

    public static XWPFDocument test(String dest) {

        FileOutputStream out = null; // 创建输出流
        XWPFDocument doc = new XWPFDocument();
        //信纸
//        WordTemplate.setDocSize(doc, "15840", "12240", STPageOrientation.Enum.forString("landscape"));
        //A4大小和横向
        XWPFWordUtils.setDocSize(doc, "16837", "11905", STPageOrientation.Enum.forString("landscape"));
        //设置纸张边距
        XWPFWordUtils.setDocMargin(doc, "1000", "500", "1000", "500");
        //设置页眉
//        XWPFWordUtils.createDefaultHeader(doc, "联谊大学学生成绩单");
        try {
            setScoreInfo(doc);
            out = new FileOutputStream(new File(dest));
            doc.write(out);
            return doc;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {

                    e.printStackTrace();
                }
            }
        }
        return doc;
    }

    private static void setScoreInfo(XWPFDocument doc) {
        for (int i = 0; i < Constants.CNT; i++) {

            XWPFWordUtils.getRun(doc, "联谊科技大学学生成绩单", 12, true);
            //设置学生基本信息
            setStudentInfo(doc);

            XWPFWordUtils.getRun(doc, ".", 1, false);

            //填充成绩信息
            setScore(doc);

            XWPFParagraph p1 = doc.createParagraph(); // 创建段落
            p1.setPageBreak(true); //设置分页符
        }
    }

    private static void setScore(XWPFDocument doc) {
        //创建表格
        XWPFTable table = XWPFWordUtils.getTable(doc, Constants.ROW_CNT, TitleEnum.values().length * Constants.BIG_COL_CNT,false);
        XWPFTableRow tableRow = table.getRow(0);
        List<XWPFTableCell> cellList = tableRow.getTableCells();
        for (int j = 0; j < cellList.size(); j++) {
            XWPFTableCell cell = cellList.get(j);
            //设置内容水平居中及列宽
            XWPFWordUtils.fillTableTitle(cell, XWPFWidthEnum.getByCode(j % 5).getMsg(), TitleEnum.getByCode(j % 5).getMsg(), 8, true);
        }
        for (int i = 1; i < Constants.ROW_CNT; i++) {
            for (int j = 0; j < Constants.WITHS.length; j++) {
                XWPFWordUtils.fillCell(table, i, j, "测试", 8);
            }
        }
        //合并单元格
        XWPFWordUtils.mergeCellsHorizontal( table,  Constants.ROW_CNT-1,  10,  14);
        XWPFWordUtils.fillCell(table, Constants.ROW_CNT-1, 10, "测的晚餐试", 8,true);
        XWPFWordUtils.fillCell(table, Constants.ROW_CNT-1, 10, "测的晚435575餐试", 8,true);

    }

    private static void setStudentInfo(XWPFDocument doc) {
        int titleFontSize = 10;
        //创建表格
        XWPFTable table0 = XWPFWordUtils.getTable(doc, 2, 4,true);
        XWPFTableRow tableRow0 = table0.getRow(0);
        List<XWPFTableCell> cellList0 = tableRow0.getTableCells();
        XWPFWordUtils.fillCellLeft(cellList0.get(0), "2400", "学院：汽车技术与服务学院", titleFontSize, true);
        XWPFWordUtils.fillCellLeft(cellList0.get(1), "2400", "专业：2018奔驰现代学徒制冠名班", titleFontSize, true);
        XWPFWordUtils.fillCellLeft(cellList0.get(2), "2400", "班级：QC汽车奔驰1804", titleFontSize, true);
        XWPFWordUtils.fillCellLeft(cellList0.get(3), "2400", "学制：3", titleFontSize, true);

        XWPFTableRow tableRow1 = table0.getRow(1);
        List<XWPFTableCell> cellList1 = tableRow1.getTableCells();
        XWPFWordUtils.fillCellLeft(cellList1.get(0), "7200", "姓名：郑攀", titleFontSize, true);
        XWPFWordUtils.fillCellLeft(cellList1.get(1), "7200", "学号：2018108QC0220", titleFontSize, true);
        XWPFWordUtils.fillCellLeft(cellList1.get(2), "7200", "入校时间：2019-09-10", titleFontSize, true);
        XWPFWordUtils.fillCellLeft(cellList1.get(3), "7200", "培养方式：全日制", titleFontSize, true);
    }
}
