package dev.c20.commons.excel.tools;


import java.util.LinkedList;
import java.util.List;

public class JsonExcelReport {
    private List<ReportParam> params = new LinkedList<>();
    private String resourceJson;
    private String outputName;

    public String getOutputName() {
        return outputName;
    }

    public JsonExcelReport setOutputName(String outputName) {
        this.outputName = outputName;
        return this;
    }

    public List<ReportParam> getParams() {
        return params;
    }

    public JsonExcelReport setParams(List<ReportParam> params) {
        this.params = params;
        return this;
    }

    public String getResourceJson() {
        return resourceJson;
    }

    public JsonExcelReport setResourceJson(String resourceJson) {
        this.resourceJson = resourceJson;
        return this;
    }

}
