package dev.c20.commons.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import dev.c20.commons.excel.helpers.CrossTabOneLevel;
import dev.c20.commons.excel.helpers.CrossTabTwoLevels;
import dev.c20.commons.excel.tools.Field;
import dev.c20.commons.excel.tools.JsonExcelReport;
import dev.c20.commons.excel.tools.RecordSet;
import dev.c20.commons.excel.tools.ReportParam;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class ExcelGeneratorFromJSON extends ExcelGenerator {

    Map<String, Object> report = null;
    JsonExcelReport jsonExcelReport;
    long startReport;
    public void createReport(JsonExcelReport jsonExcelReport) throws Exception {
        startReport = System.currentTimeMillis();

        configure(jsonExcelReport);


        createRecordSets();

        finishExcel();


        System.out.println( "finished in " + ( System.currentTimeMillis() - startReport) + " ms");

    }


    public ExcelGeneratorFromJSON configure( JsonExcelReport jsonExcelReport) throws JsonProcessingException {
        this.jsonExcelReport = jsonExcelReport;
        setPrameters(jsonExcelReport.getParams());
        String json = resourceAsString(jsonExcelReport.getResourceJson() );

        ObjectMapper mapper = new ObjectMapper();
        report = mapper.readValue(json, Map.class);

        System.out.println( json );

        configureStyles();
        configureColumnsWidths();
        writeFixedCells();
        writeParams();
        writeRecordSets();
        return this;
    }

    public ExcelGeneratorFromJSON configureStyles() {

        List<Map<String,Object>> styles = (List<Map<String,Object>>)report.get("styles");

        for( Map<String,Object> style : styles ) {
            //{ "name":  "header-rechazada", "color":  "0x000000", "backgroundColor":  "251,235,247", "horizontalAlignment":  "CENTER", "bold":  false }
            String name = (String)style.get("name");
            String color = (String)style.get("color");
            String backgroundColor = (String)style.get("backgroundColor");
            String horizontalAlignment = (String)style.get("horizontalAlignment");
            Boolean bold = (Boolean)style.get("bold");
            String borderBottom = (String)style.get("borderBottom");
            String borderTop = (String)style.get("borderTop");
            String borderLeft = (String)style.get("borderLeft");
            String borderRight = (String)style.get("borderRight");


            XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.createCellStyle();
            Font font = workbook.createFont();
            XSSFFont xssfFont = (XSSFFont)font;
            if( color != null ) {

                XSSFColor colorFore = new XSSFColor(getColor(color), null);
                xssfFont.setColor(colorFore);
            }
            if( bold != null)
                xssfFont.setBold(bold);

            cellStyle.setFont(font);

            if( backgroundColor != null ) {
                XSSFColor colorBack = new XSSFColor(getColor(backgroundColor), null);
                cellStyle.setFillForegroundColor(colorBack);
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            if( horizontalAlignment != null )
                cellStyle.setAlignment(HorizontalAlignment.valueOf(horizontalAlignment));

            if( borderTop != null )
                cellStyle.setBorderTop(BorderStyle.valueOf(borderTop));

            if( borderBottom != null )
                cellStyle.setBorderBottom(BorderStyle.valueOf(borderBottom));

            if( borderLeft != null )
                cellStyle.setBorderLeft(BorderStyle.valueOf(borderLeft));

            if( borderRight != null )
                cellStyle.setBorderRight(BorderStyle.valueOf(borderRight));

            this.styles.put(name,cellStyle);



        }

        return this;
    }

    public ExcelGeneratorFromJSON configureColumnsWidths() {

        List<Map<String,Object>> widths = (List<Map<String,Object>>)report.get("columnsWidth");

        for( Map<String,Object> columnWidth : widths ) {
            String fromCellRef = (String)columnWidth.get("fromCell");
            String toCellRef = (String)columnWidth.get("toCell");
            float width = 0;
            if( columnWidth.get("width") instanceof  Integer ) {
                width = Float.valueOf( "" + columnWidth.get("width") );
            } else {
                width = (float)((double)columnWidth.get("width"));
            }


            if( toCellRef == null ) {
                setColumnWidth( fromCellRef, width );
            } else {
                setColumnWidth( fromCellRef, toCellRef,  width );
            }

        }

        return this;
    }

    public ExcelGeneratorFromJSON writeFixedCells() {
        List<Map<String,Object>> cells = (List<Map<String,Object>>)report.get("fixedCells");

        for( Map<String,Object> cell : cells ) {
            String cellRef = (String)cell.get("cell");
            String value = (String)cell.get("value");
            String style =(String) cell.get("style");

            if( cell.get("verticalValues") != null ) {
                List<String> values = (List<String>)cell.get("verticalValues");
                int row = (Integer)cell.get("row") -1;
                int col = CellReference.convertColStringToIndex(cellRef);

                for( String val : values ) {
                    setCell(row++,col,val,style);
                }
                continue;
            }

            if( cell.get("horizontalValues") != null) {
                List<String> values = (List<String>)cell.get("horizontalValues");
                int row = (Integer)cell.get("row") -1;
                int col = CellReference.convertColStringToIndex(cellRef);

                for( String val : values ) {
                    setCell(row,col++,val,style);
                }
                continue;
            }

            setCell(cellRef,value,style);

        }

        return this;
    }

    public Map<String,Object> findParam(String name)  {
        Map<String,Object> result = null;

        Map<String,Object> recordSets = (Map<String,Object>)report.get("sqlRecordSets");

        for( String recordSetName : recordSets.keySet() ) {
            Map<String,Object> recordSet = (Map<String,Object>)recordSets.get(recordSetName);
            List<Map<String, Object>> params = (List<Map<String, Object>>) recordSet.get("params");
            for( Map<String, Object> param : params ) {
                if( ((String)param.get("name")).equalsIgnoreCase(name) ) {
                    return param;
                }
            }
        }


        return result;
    }

    public ExcelGenerator writeParams() {
        for(ReportParam rp : parameters ){
            Map<String,Object> cfgParam = findParam(rp.getName());
            if( cfgParam != null ) {
                setCell((String)cfgParam.get("value"), rp.getValue(), (String)cfgParam.get("style"));
            } else {
                setCell(rp.getCell(), rp.getValue(), rp.getStyle());
            }
        }
        return this;
    }

    public ExcelGeneratorFromJSON writeRecordSets() {
        writeRecordSet("none");
        return this;
    }

    public ExcelGeneratorFromJSON writeRecordSet(String name) {
        return this;
    }

    public Color getColor(String color) {
        if( color == null ) {
            return null;
        }
        if( color.startsWith("0x") ) {
            return new Color( Integer.decode(color) );
        } else  {
            String[] rgb = color.split(",");
            String hex = String.format("%02x%02x%02x", Integer.parseInt( rgb[0] ), Integer.parseInt( rgb[1] ), Integer.parseInt( rgb[2] ));
            return new Color(Integer.decode("0x"+hex));
        }
    }


    Connection conn;
    public ExcelGeneratorFromJSON setConnection(Connection  conn) {
        this.conn = conn;
        return this;
    }

    public void runCustomRecordSet( Map<String,Object>  recordSetCfg ) {

    }

    public ExcelGeneratorFromJSON createRecordSets() {
        List<Map<String,Object>> recordSets = (List<Map<String,Object>>)report.get("recordSets");

        for( Map<String,Object> recordSetCfg : recordSets ) {
            Integer index = (Integer)recordSetCfg.get("index");
            String name = (String)recordSetCfg.get("name");
            String type = (String)recordSetCfg.get("type");
            String recordSet = (String)recordSetCfg.get("recordSet");

            type = type == null ? "recordset" : type;

            if( type.equalsIgnoreCase("custom") ) {
                runCustomRecordSet(recordSetCfg);
                continue;
            } else if( type.equalsIgnoreCase("memory")) {
                System.out.println("Put in memory:" + recordSet);
                this.recordSets.put( name, sqlRunRecordSet(recordSet, false,null) );
                continue;
            } else if( type.equalsIgnoreCase("crosstab-two-levels")) {
                new CrossTabTwoLevels(this,recordSetCfg);
                continue;
            } else if( type.equalsIgnoreCase("crosstab-one-level")) {
                new CrossTabOneLevel(this,recordSetCfg);
                continue;
            }

            String direction = (String)recordSetCfg.get("direction");
            Integer fromRow = (Integer)recordSetCfg.get("fromRow");
            Integer fromCol = (Integer)recordSetCfg.get("fromCol");
            Integer toRow = (Integer)recordSetCfg.get("toRow");
            Integer toCol = (Integer)recordSetCfg.get("toCol");
            Boolean mergeCells = (Boolean)recordSetCfg.get("mergeCells");
            Integer colWidth = (Integer)recordSetCfg.get("colWidth");
            Boolean haveStyle = (Boolean)recordSetCfg.get("recordSetWithStyles");
            String defaultStyle = (String)recordSetCfg.get("recordSet");
            haveStyle = haveStyle == null ? true : haveStyle;
            RecordSet rs = sqlRunRecordSet(recordSet, haveStyle,defaultStyle);
            if (direction.equalsIgnoreCase("down")) {
                writeVerticalRecordSet(fromRow, fromCol, rs);
            } else {
                writeHorizontalRecordSet(fromRow, fromCol, rs, colWidth, mergeCells);
            }
        }
        return this;
    }

    public RecordSet sqlRunRecordSet(String recordSetName ) {
        return sqlRunRecordSet(recordSetName,true,null);
    }
    public RecordSet sqlRunRecordSet(String recordSetName, boolean haveStyle, String style ) {
        System.out.println("Run RecordSet:" + recordSetName);

        Map<String,Map<String,Object>> sqlRs = (Map<String,Map<String,Object>>)report.get( "sqlRecordSets");
        Map<String,Object> sqlCfg = (Map<String,Object>)sqlRs.get(recordSetName);
        List<String> sqlList = (List<String>)sqlCfg.get("sql");
        int columnsFrom = (Integer)sqlCfg.get("columnsFrom");
        System.out.println( recordSetName );

        String sql = "";
        String optionalParams = "";



        List<Map<String,Object>> params = (List<Map<String,Object>>)sqlCfg.get("params");
        /*
        System.out.println(sql);
        System.out.println("Columns from " + columnsFrom);
        */
        try {

            Map<String,String> readedParams = new HashMap<>();

            for( Map<String,Object> param : params ) {
                Cell cell = getCellWithoutCreate((String)param.get("value"));

                if( cell != null ) {
                    if( param.get("optional") != null && (boolean) param.get("optional")) {
                        optionalParams += param.get("sql") + "\n";
                    }
                }
            }


            for( String line : sqlList ) {
                if( line.equalsIgnoreCase( "/* optional parameters*/") )  {
                    if( optionalParams != null ) {
                        sql += optionalParams + "\n";
                    }
                } else {
                    sql += line + "\n";
                }
            }

            System.out.println("Optional paramers found:" +optionalParams);

            PreparedStatement ps = conn.prepareStatement(sql);
            int paramNum = 1;
            for( Map<String,Object> param : params ) {
                Cell cell = getCellWithoutCreate((String)param.get("value"));

                System.out.print( "Param (" + paramNum+ ") Name:" + param.get("name") + " from cell:" + param.get("value") + " as ");
                if( cell == null ) {
                    paramNum++;
                    continue;
                }
                if( cell.getCellType() == CellType.BOOLEAN ) {
                    System.out.println("boolean");
                    ps.setBoolean( paramNum, cell.getBooleanCellValue() );
                } else if( cell.getCellType() == CellType.NUMERIC ) {
                    System.out.println("numeric");
                    ps.setDouble(paramNum, cell.getNumericCellValue());
                } else if( cell.getCellType() == CellType.STRING ) {
                    System.out.println("string");
                    ps.setString(paramNum, cell.getStringCellValue());

                } else {
                    System.out.println(" type not found");
                }
                paramNum++;
            }
            ResultSet rs = ps.executeQuery();

            RecordSet result = new RecordSet();
            result.setColCount(ps.getMetaData().getColumnCount() - columnsFrom);
            while( rs.next() ) {

                List<Field> record = new LinkedList<>();
                if( haveStyle ) {
                    for (int col = columnsFrom; col <= ps.getMetaData().getColumnCount(); col += 2) {
                        Field field = new Field(rs.getString(col), rs.getObject(col + 1));
                        record.add(field);
                    }
                } else {
                    for (int col = columnsFrom; col <= ps.getMetaData().getColumnCount(); col ++) {
                        Field field = new Field(style, rs.getObject(col));
                        record.add(field);
                    }
                }
                result.getRecords().add(record);
            }

            result.setRowCount(result.getRecords().size());


            rs.close();
            ps.close();

            return result;
        } catch ( Exception ex ) {
            System.out.println(sql);
            System.out.println("Columns from " + columnsFrom);
            ex.printStackTrace();
        }

        return null;
    }


}
