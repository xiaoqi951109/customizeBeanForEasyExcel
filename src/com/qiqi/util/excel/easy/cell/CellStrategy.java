package com.qiqi.util.excel.easy.cell;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 样式处理
 */
public class CellStrategy implements CellWriteHandler {

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, org.apache.poi.ss.usermodel.Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (cellData.getData() instanceof Cell) {
            Cell buCell = (Cell) cellData.getData();
            if (buCell.row > 1 || buCell.col > 1) {
                writeSheetHolder.getSheet().addMergedRegionUnsafe(
                        new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex() + buCell.row - 1, cell.getColumnIndex(), cell.getColumnIndex() + buCell.col - 1));
            }
            CellStyle style = writeSheetHolder.getSheet().getWorkbook().createCellStyle();
            short backgroundColor = buCell.bc;
            if (backgroundColor > 0) {
                style.setFillForegroundColor(buCell.bc);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            HorizontalAlignment alignment = buCell.a;
            if (null != alignment) {
                style.setAlignment(alignment);
            }
            cell.setCellStyle(style);

        }
    }

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, org.apache.poi.ss.usermodel.Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, org.apache.poi.ss.usermodel.Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
    }

}
