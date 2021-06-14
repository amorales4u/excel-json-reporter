package dev.c20.commons.excel.tools;


import java.util.LinkedList;
import java.util.List;

public class RecordSet {

    private int rowCount;
    private int  colCount;
    List<List<Field>> records = new LinkedList<>();

    public int getRowCount() {
        return rowCount;
    }

    public Field get( int row, int col) {
        return records.get(row).get(col);
    }

    public RecordSet setRowCount(int rowCount) {
        this.rowCount = rowCount;
        return this;
    }

    public int getColCount() {
        return colCount;
    }

    public RecordSet setColCount(int colCount) {
        this.colCount = colCount;
        return this;
    }

    public List<List<Field>> getRecords() {
        return records;
    }

    public RecordSet setRecords(List<List<Field>> records) {
        this.records = records;
        return this;
    }
}
