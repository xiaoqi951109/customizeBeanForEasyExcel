package com.qiqi.util.excel.easy.cell;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.ArrayList;
import java.util.List;

/**
 * Cell 封装对象
 * 可以将所有的类和JSON等数据进行统一转换和处理
 * 属性命名简短化是考虑到可能对大数据的资源占用可能有一定好处
 *
 * @Author XiaoQi
 * @Time 2021/3/18
 */
public class Cell {

    public Cell(Object v) {
        this.v = v;
        this.row = 1;
        this.col = 1;
    }

    String t;// t

    Object v;// value

    short bc;// background color

    short fc;// force color

    int fs;//font size

    int row;//row count

    int col;//column count

    HorizontalAlignment a;//align

    public Cell t(String t) {
        this.t = t;
        return this;
    }

    public Cell v(Object v) {
        this.v = v;
        return this;
    }

    public short bc() {
        return bc;
    }

    public Cell bc(short bc) {
        this.bc = bc;
        return this;
    }

    public Cell fc(short fc) {
        this.fc = fc;
        return this;
    }

    public Cell fs(int fs) {
        this.fs = fs;
        return this;
    }

    public Cell a(HorizontalAlignment a) {
        this.a = a;
        return this;
    }

    public Cell row(int row) {
        this.row = row;
        return this;
    }

    public Cell col(int col) {
        this.col = col;
        return this;
    }


    public List<Cell> get() {
        if (this.col < 2) {
            this.col = 1;
        }
        List<Cell> cellList = new ArrayList(this.col);
        cellList.add(this);
        for (int i = 1; i < this.col; i++) {
            cellList.add(null);
        }
        return cellList;
    }

}
