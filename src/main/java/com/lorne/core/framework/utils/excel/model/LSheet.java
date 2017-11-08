package com.lorne.core.framework.utils.excel.model;

import java.util.List;

public class LSheet {

    private String sheetName;

    private List<LRow> rows;

    public String getSheetName() {
        return sheetName;
    }
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<LRow> getRows() {
        return rows;
    }

    public void setRows(List<LRow> rows) {
        this.rows = rows;
    }
}
