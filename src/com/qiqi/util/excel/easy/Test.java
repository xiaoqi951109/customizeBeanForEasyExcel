package com.qiqi.util.excel.easy;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.qiqi.util.excel.easy.cell.CellConverter;
import com.qiqi.util.excel.easy.cell.CellStrategy;
import com.qiqi.util.excel.easy.cell.RowFactory;
import com.qiqi.util.excel.easy.cell.SheetFactory;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;

public class Test {

    public static void main(String[] args) {
        String fileName = "\\xxxx\\simpleWrite.xlsx";
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(fileName).build();
//            List data = demoDataList1();//跨行跨列的坐标演示
            List data = demoDataList2();//跨行跨列的处理
            WriteSheet writeSheet = EasyExcel.writerSheet("sheet1")
                    .registerConverter(new CellConverter())  //数据转换注册
                    .registerWriteHandler(new CellStrategy())  //Cell处理钩子
                    .build();
            excelWriter.write(data, writeSheet);
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 通过 x.y的方式感知跨行跨列的处理
     *
     * @return
     */
    private static List demoDataList1() {
        SheetFactory sheetFactory = new SheetFactory();

        RowFactory factory1 = new RowFactory();
        factory1.num("1.1");
        factory1.str("1.2");
        factory1.str("1.3-5").col(3).a(HorizontalAlignment.CENTER);
        factory1.str("1.6");
        factory1.num("1.7").bc(IndexedColors.RED.getIndex());
        sheetFactory.add(factory1);

        RowFactory factory2 = new RowFactory();
        factory2.num(2.1).row(2).col(3).a(HorizontalAlignment.RIGHT);
        factory2.str("2.4");
        factory2.str("2.5");
        factory2.str("2.6").bc(IndexedColors.CORAL.getIndex());
        sheetFactory.add(factory2);

        RowFactory factory3 = new RowFactory();
        factory3.num("3.4");
        factory3.str("3.5").bc(IndexedColors.BLUE.getIndex());
        sheetFactory.add(factory3);

        RowFactory factory4 = new RowFactory();
        factory4.str("4.1");
        factory4.str("4.2-3").a(HorizontalAlignment.CENTER).col(2);
        factory4.num("4.4").bc(IndexedColors.AQUA.getIndex());
        sheetFactory.add(factory4);

        return sheetFactory.get();
    }


    /**
     * 跨行列实际的excel模板
     *
     * @return
     */
    private static List demoDataList2() {
        SheetFactory sheetFactory = new SheetFactory();

        RowFactory factory0 = new RowFactory();
        factory0.str("订单汇总模板").row(3).col(4).bc(IndexedColors.GREY_40_PERCENT.getIndex()).a(HorizontalAlignment.CENTER);
        sheetFactory.add(factory0, 3);

        RowFactory factory1 = new RowFactory();
        factory1.str("客户名称").a(HorizontalAlignment.CENTER);
        factory1.str("月份");
        factory1.str("金额");
        factory1.str("合计");
        sheetFactory.add(factory1);

        RowFactory factory2 = new RowFactory();
        factory2.str("老魏批发部").row(6).a(HorizontalAlignment.CENTER);
        factory2.str("3月");
        factory2.num("8.0");
        factory2.num("30.0").row(6).a(HorizontalAlignment.CENTER);
        sheetFactory.add(factory2);

        RowFactory factory3 = new RowFactory();
        factory3.str("4月");
        factory3.num("3.0");
        sheetFactory.add(factory3);

        RowFactory factory32 = new RowFactory();
        factory32.str("5月");
        factory32.num("6.0");
        sheetFactory.add(factory32);

        RowFactory factory33 = new RowFactory();
        factory33.str("6月");
        factory33.num("3.0");
        sheetFactory.add(factory33);

        RowFactory factory34 = new RowFactory();
        factory34.str("7月");
        factory34.num("6.0");
        sheetFactory.add(factory34);


        RowFactory factory35 = new RowFactory();
        factory35.str("8月");
        factory35.num("4.0");
        sheetFactory.add(factory35);


        RowFactory factory5 = new RowFactory();
        factory5.str("小张便利店").row(2).a(HorizontalAlignment.CENTER);
        factory5.str("3月");
        factory5.num("35.0");
        factory5.num("38.0").row(2).a(HorizontalAlignment.CENTER);
        sheetFactory.add(factory5);


        RowFactory factory51 = new RowFactory();
        factory51.str("4月");
        factory51.num("3.0");
        sheetFactory.add(factory51);

        return sheetFactory.get();
    }
}
