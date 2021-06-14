package dev.c20.commons.excel.tools;


public class Field {
    String style;
    Object value;

    public Field( String style, Object value) {
        this.style = style;
        this.value = value;
    }

    public String getStyle() {
        return style;
    }

    public Field setStyle(String style) {
        this.style = style;
        return this;
    }

    public Object getValue() {
        return value;
    }

    public Field setValue(Object value) {
        this.value = value;
        return this;
    }
}
