package com.qiqi.util.excel.easy.cell;

import java.util.ArrayList;
import java.util.List;

/**
 * Row Set
 *
 * @Author XiaoQi
 * @Time 2021/3/18
 */
public class RowFactory {

    // cell in row 用于记录代码插入进Row的数据，因为涉及到上一行跨列的影响，所以和最终插入的cellList可能有差异
    List<Cell> cellList = new ArrayList<>();
    // cell use row state 用于记录每一个cell位 跨列情况，往下跨N列就存N，然后下一列自动继承上一列N-1的数值
    List<Integer> stateList = new ArrayList<>();

    /**
     * 根据上一行的状态集合和当前记录的cell集合动态构建最终插入的cell集合
     * stateList 往下跨N列就存N，然后下一列自动继承上一列N-1的数值 ,如果>0表名新的一行目标位置的cell被上行跨列占用了，就把state-1进行继续传递，以达到跨行跨列的效果
     *
     * @param lastRowState
     * @return
     */
    public List<Cell> list(List<Integer> lastRowState) {
        List<Cell> list = new ArrayList<>();
        int currentIndex = 0;
        for (int i = 0; i < lastRowState.size(); i++) {
            if (lastRowState.get(i) > 0) {
                stateList.add(lastRowState.get(i) - 1);
                list.add(null);
            } else {
                if (cellList.size() > currentIndex) {
                    convertItem(cellList.get(currentIndex), list);
                    currentIndex++;
                }
            }
        }
        for (int i = currentIndex; i < cellList.size(); i++) {
            convertItem(cellList.get(i), list);
        }
        return list;
    }

    /**
     * cell和state的动态构建
     *
     * @param cell
     * @param list
     */
    private void convertItem(Cell cell, List<Cell> list) {
        if (null != cell) {
            for (int j = 0; j < cell.col; j++) {
                stateList.add(cell.row - 1);
            }
            list.addAll(cell.get());
        } else {
            stateList.add(0);
            list.add(null);
        }
    }

    /**
     * 插入Number类型的数据
     *
     * @param v
     * @return
     */
    public Cell num(Object v) {
        Cell cell = new Cell(v).t("double");
        cellList.add(cell);
        return cell;
    }

    /**
     * 插入字符类型的数据
     *
     * @param v
     * @return
     */
    public Cell str(Object v) {
        Cell cell = new Cell(v);
        cellList.add(cell);
        return cell;
    }

    /**
     * 获取当前行的使用状态
     *
     * @return
     */
    public List<Integer> getStateList() {
        return stateList;
    }
}
