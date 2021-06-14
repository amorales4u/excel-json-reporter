package dev.c20.commons.excel.helpers;


import java.util.LinkedList;
import java.util.List;

public class CrossCell {
    CrossCellType type;
    Object value;
    String style;

    List<Object> params = new LinkedList<>();
    public CrossCell() {

    }
    public CrossCell(CrossCellType type, Object value ) {
        this.type = type;
        this.value = value;
    }
    public CrossCell(CrossCellType type, Object value, String style ) {
        this.type = type;
        this.value = value;
        this.style = style;
    }
}
