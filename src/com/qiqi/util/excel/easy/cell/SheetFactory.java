package com.qiqi.util.excel.easy.cell;

import java.util.ArrayList;
import java.util.List;

/**
 * Sheet Set
 *
 * @Author XiaoQi
 * @Time 2021/3/18
 */
public class SheetFactory {

    List<RowFactory> rowList = new ArrayList<>();

    public void add(RowFactory rowFactory) {
        add(rowFactory, 1);
    }

    public void add(RowFactory rowFactory, int rowLength) {
        rowList.add(rowFactory);
        fill(rowLength - 1);
    }

    public void fill(int size) {
        for (int i = 0; i < size; i++) {
            rowList.add(new RowFactory());
        }
    }

    public List<List<Cell>> get() {
        List<List<Cell>> list = new ArrayList<>();
        for (int i = 0; i < rowList.size(); i++) {
            list.add(rowList.get(i).list(getLastRowState(i - 1)));
        }
        return list;
    }

    private List<Integer> getLastRowState(int index) {
        if (index > -1) {
            RowFactory rowFactory = rowList.get(index);
            return rowFactory.getStateList();
        }
        return new ArrayList<>();
    }


}
