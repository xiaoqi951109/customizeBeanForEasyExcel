package com.qiqi.util.excel.easy.cell;

import java.util.ArrayList;
import java.util.List;

/**
 *  Row Set
 *
 * @Author XiaoQi
 * @Time 2021/3/18
 */
public class RowFactory {

    // cell in row
    List<Cell> cellList = new ArrayList<>();
    // cell use row state
    List<Integer> stateList = new ArrayList<>();

    public List<Cell> list(List<Integer> lastRowState) {
        List<Cell> list = new ArrayList<>();
        Cell cell;
        int currentIndex = 0;
        for (int i = 0; i < lastRowState.size(); i++) {
            if (lastRowState.get(i) > 0) {
                stateList.add(lastRowState.get(i) - 1);
                list.add(null);
            } else {
                if (cellList.size() > currentIndex) {
                    cell = cellList.get(currentIndex);
                    if (null != cell) {
                        for (int j = 0; j < cell.col; j++) {
                            stateList.add(cell.row - 1);
                        }
                        list.addAll(cell.get());
                    } else {
                        stateList.add(0);
                        list.add(cell);
                    }
                    currentIndex++;
                }
            }
        }
        for (int i = currentIndex; i < cellList.size(); i++) {
            cell = cellList.get(i);
            if (null != cell) {
                for (int j = 0; j < cell.col; j++) {
                    stateList.add(cell.row - 1);
                }
                list.addAll(cell.get());
            } else {
                stateList.add(0);
                list.add(cell);
            }
        }
        return list;
    }

    public Cell num(Object v) {
        Cell cell = new Cell(v).t("double");
        cellList.add(cell);
        return cell;
    }

    public Cell str(Object v) {
        Cell cell = new Cell(v);
        cellList.add(cell);
        return cell;
    }

    public List<Integer> getStateList() {
        return stateList;
    }
}
