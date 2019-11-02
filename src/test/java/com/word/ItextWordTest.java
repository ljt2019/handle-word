package com.word;

import com.lowagie.text.*;
import com.lowagie.text.rtf.RtfWriter2;
import com.tiger.constant.Constants;
import com.tiger.enums.TitleEnum;
import com.tiger.itext.ItextWordUtils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * 该方式生成的文件比较大，建议用poi
 *
 * @description:
 * @author: tiger
 * @create: 2019-11-02 14:58
 */
public class ItextWordTest {

    /**
     * @param args
     */
    public static void main(String[] args) {
        long start = System.currentTimeMillis();
        drawDoc();
        long end = System.currentTimeMillis();
        System.out.println("输出 Itext word成功！");
        System.out.println(end - start);

    }

    /**
     * 绘制表格
     */
    public static void drawDoc() {
        boolean type = false;
        // 1、创建word文档,并设置纸张的大小和方向
        Document doc;
        if (type) {
            doc = new Document(PageSize.A4.rotate());
        } else {
            doc = new Document(PageSize.A4);
        }

        try {
            //2、创建doc实例对象
            RtfWriter2.getInstance(doc,
                    new FileOutputStream("C:/AAAAA/itext.doc"));
//            RtfWriter2.getInstance(doc, outputStream);
            //3、打开doc对象
            doc.open();
            //设置纸张方向
            doc.setPageSize(PageSize.A4.rotate());
            //4、设置页面边距(左右上下)
            doc.setMargins(15, 15, 15, 2);

            for (int i = 0; i < Constants.CNT; i++) {
                // TODO
                int style = Font.NORMAL;
                int alignment = Element.ALIGN_CENTER;
                Table table = ItextWordUtils.getTable(100, Constants.ROW_CNT, Constants.WITHS, 5, alignment);
                for (int j = 0; j < 15; j++) {
                    ItextWordUtils.fillCell(table, TitleEnum.getByCode(j % 5).getMsg(), 7, style, alignment);
                }
                //加入文档
                doc.add(table);
                doc.newPage();
            }
            doc.close();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
