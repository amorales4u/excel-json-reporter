package dev.c20.commons.excel.helpers;


public enum CrossCellType {
    VALUE ( "V[ %s ]"),
    FORMULA ( "F[ %s ]" ),
    READ ( "R[ %s ]"),
    EMPTY ( "E[ %s ]");

    private Object value;

    private CrossCellType(Object value) {
        this.value = value;
    }

    public Object getValue() {
        return this.value;
    }


}
