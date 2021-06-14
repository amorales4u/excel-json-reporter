package dev.c20.commons.excel.tools;


public class ReportParam {
    String name;
    String cell;
    String style;
    Object value;

    public Object getValue() {
        return value;
    }

    public ReportParam setValue(Object value) {
        this.value = value;
        return this;
    }


    public String getName() {
        return name;
    }

    public ReportParam setName(String name) {
        this.name = name;
        return this;
    }

    public String getCell() {
        return cell;
    }

    public ReportParam setCell(String cell) {
        this.cell = cell;
        return this;
    }

    public String getStyle() {
        return style;
    }

    public ReportParam setStyle(String style) {
        this.style = style;
        return this;
    }


}
