package com.qiqi.util.excel.easy.cell;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.math.BigDecimal;

/**
 * Object to cellData
 *
 * @Author XiaoQi
 * @Time 2021/3/18
 */
public class CellConverter implements Converter {
    public Class supportJavaTypeKey() {
        return Cell.class;
    }

    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    public Object convertToJavaData(CellData cellData, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        return null;
    }

    public CellData convertToExcelData(Object o, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        CellData<Cell> cellData = new CellData<>();
        Cell cell = (Cell) o;
        if (null == cell) {
            return cellData;
        }
        if (null != cell.t && null != cell.v && cell.t.equals("double")) {
            //double value
            if (!(cell.v instanceof BigDecimal)) {
                cellData.setNumberValue(BigDecimal.valueOf(Double.parseDouble(String.valueOf(cell.v))));
            } else {
                cellData.setNumberValue((BigDecimal) cell.v);
            }
            cellData.setType(CellDataTypeEnum.NUMBER);
        } else {
            //string value
            cellData.setStringValue(String.valueOf(cell.v));
            cellData.setType(CellDataTypeEnum.STRING);
        }
        cellData.setData(cell);
        return cellData;
    }
}
