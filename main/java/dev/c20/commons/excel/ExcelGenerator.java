package dev.c20.commons.excel;

import dev.c20.commons.excel.tools.Field;
import dev.c20.commons.excel.tools.RecordSet;
import dev.c20.commons.excel.tools.ReportParam;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;
import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.*;

public class ExcelGenerator {

    public ExcelGenerator writeHorizontalRecordSet(int startFromRow, int startFromCol, RecordSet recordSet, int colWidth, boolean merge ) {
        List<List<Field>> data = recordSet.getRecords();
        for( int seg = 0; seg < data.size(); seg ++ ) {
            int startCol =  startFromCol + ( seg  * 5 );
            int startRow = startFromRow;
            List<Field> record = data.get(seg);
            if( merge ) {
                mergeRegion(startCol, startRow, startCol + colWidth - 1, startRow);
            }
            for( Field field : record ) {
                setCell(startRow, startCol++, field.getValue(), field.getStyle());
            }
            //writeRecordSet( startRow + 1,startCol, TestData.getSegmento(name.getStyle(), (String)name.getValue(), seg * 2) );
        }

        return this;

    }

    public ExcelGenerator writeVerticalRecordSet(int startRow, int startCol, RecordSet recordSet) {
        if( recordSet == null )
            return this;
        List<List<Field>> data = recordSet.getRecords();

        for( int row = 0; row < data.size(); row ++  ) {
            List<Field> record = data.get(row);
            int startColRecord = startCol;
            for( int col = 0; col < record.size(); col ++  ) {
                Field field = record.get(col);
                setCell(startRow + row, startColRecord++ , field.getValue(), field.getStyle());
            }
        }
        return this;
    }


    public ExcelGenerator() {
        prepareExcel();
    }

    public ExcelGenerator mergeRegion( int fromCol, int fromRow, int toCol, int toRow ) {
        sheet.addMergedRegion(new CellRangeAddress(fromRow, toRow, fromCol, toCol));
        return this;
    }
    public ExcelGenerator mergeRegion( String regionRef ) {
        String[] cellsRef = regionRef.split(":");
        CellReference fromCellReference = new CellReference(cellsRef[0]);
        CellReference toCellReference = new CellReference(cellsRef[1]);
        sheet.addMergedRegion(new CellRangeAddress(fromCellReference.getRow(), toCellReference.getRow(), fromCellReference.getCol(), toCellReference.getCol()));
        return this;
    }

    public int getRow(String row) {
        return Integer.parseInt(row) - 1;
    }

    public int getCol(String col) {
        CellReference fromCellReference = new CellReference(col + "1");
        return fromCellReference.getCol();
    }
    /*
        public ExcelGenerator createStyles() {
            //new Color( Integer.decode("0xAA0F245C") );

            setStyle("title", new java.awt.Color(0,0,0),null, HorizontalAlignment.LEFT,true);
            setStyle("header", new java.awt.Color(255,255,255),new java.awt.Color(48,84,150), HorizontalAlignment.CENTER,false);
            setStyle("header-segmento-0", new java.awt.Color(255,255,255),new java.awt.Color(89, 128, 208), HorizontalAlignment.CENTER,false);
            setStyle("header-segmento-1", new java.awt.Color(255,255,255),new java.awt.Color(30, 161, 194), HorizontalAlignment.CENTER,false);
            setStyle("header-en-tiempo", new java.awt.Color(255,255,255),new java.awt.Color(0,176,80), HorizontalAlignment.CENTER,false);
            setStyle("header-vencida", new java.awt.Color(255,255,255),new java.awt.Color(192,0,0), HorizontalAlignment.CENTER,false);
            setStyle("header-no-atendida", new java.awt.Color(0,0,0),new java.awt.Color(255,242,204), HorizontalAlignment.CENTER,false);
            setStyle("header-rechazada", new java.awt.Color(0,0,0),new java.awt.Color(251,235,247), HorizontalAlignment.CENTER,false);
            setStyle("header-entidad", new java.awt.Color(255,255,255),new java.awt.Color(47,117,181), HorizontalAlignment.LEFT,false);
            setStyle("header-entidad-num", new java.awt.Color(255,255,255),new java.awt.Color(47,117,181), HorizontalAlignment.CENTER,false);
            setStyle("row-solicitud", new java.awt.Color(0,0,0),new java.awt.Color(255,255,255), HorizontalAlignment.LEFT,false);
            setStyle("row-solicitud-num", new java.awt.Color(0,0,0),new java.awt.Color(255,255,255), HorizontalAlignment.CENTER,false);
            return this;
        }
    */
    public SXSSFCell setCell( int rowNum, int colNum, String value ) {
        SXSSFCell cell = getCell( rowNum, colNum );

        cell.setCellValue(value);
        return cell;
    }

    public SXSSFCell setCell( String cellRef, String value ) {
        SXSSFCell cell = getCell( cellRef);

        cell.setCellValue(value);
        return cell;
    }

    public SXSSFCell setCell( String cellRef, Object value ) {
        SXSSFCell cell = getCell( cellRef);

        setCellValue(cell,value);
        return cell;
    }

    public SXSSFCell setCellValue( SXSSFCell cell, Object value) {
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Double ) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if( value instanceof BigDecimal ) {
            //cell.setCellValue( ((BigDecimal)value).toPlainString() );
            cell.setCellValue( ((BigDecimal)value).doubleValue() );
        }

        return cell;
    }

    public SXSSFCell setCell( int rowNum, int colNum, Object value, String styleName ) {
        SXSSFCell cell = getCell(rowNum, colNum);

        cell.setCellStyle(styles.get(styleName));
        setCellValue(cell,value);
        return cell;
    }

    public SXSSFCell setCellStyle( int rowNum, int colNum, String styleName ) {
        SXSSFCell cell = getCell(rowNum, colNum);
        XSSFCellStyle style = styles.get(styleName);

        cell.setCellStyle(style);
        return cell;
    }

    public SXSSFCell setFormula( int rowNum, int colNum, String formula, String styleName ) {
        SXSSFCell cell = getCell(rowNum, colNum);

        cell.setCellStyle(styles.get(styleName));
        cell.setCellFormula(formula);

        return cell;
    }


    public SXSSFCell setCell( String cellRef, Object value, String styleName ) {
        SXSSFCell cell = getCell( cellRef);

        cell.setCellStyle(styles.get(styleName));
        setCellValue(cell,value);
        return cell;
    }


    public ExcelGenerator setColumnWidth( String fromCellRef, float width ) {
        setColumnWidth( fromCellRef, fromCellRef, width);
        return this;
    }

    public ExcelGenerator setColumnWidth( String fromCellRef, String toCellRef, float width ) {
        CellReference fromCellReference = new CellReference(fromCellRef);
        CellReference toCellReference = new CellReference(toCellRef);

        for( int idx = fromCellReference.getCol(); idx <= toCellReference.getCol(); idx ++ ) {
            sheet.setColumnWidth(idx, getExcelWidth(width) );
        }
        return this;
    }

    public int getExcelWidth(float widthExcel ) {
        int width256 = (int)Math.floor((widthExcel * Units.DEFAULT_CHARACTER_WIDTH + 5) / Units.DEFAULT_CHARACTER_WIDTH * 256);
        return width256;
    }

    public SXSSFCell getCell( int rowNum, int colNum) {
        Row row = sheet.getRow(rowNum);
        if( row == null ) {
            row = sheet.createRow(rowNum);
        }
        Cell cell = row.getCell(colNum);
        if( cell == null ) {
            cell = row.createCell(colNum);
        }
        return (SXSSFCell)cell;
    }

    public CellReference getCellReference( String cellRef ) {
        return new CellReference(cellRef);
    }

    public SXSSFCell getCellWithoutCreate( String cellRef) {
        CellReference cellReference = new CellReference(cellRef);
        Row row = sheet.getRow(cellReference.getRow());
        if( row == null ) {
            return null;
        }
        Cell cell = row.getCell(cellReference.getCol());
        if( cell == null ) {
            return null;
        }
        return (SXSSFCell)cell;
    }

    public SXSSFCell getCell( String cellRef) {
        CellReference cellReference = new CellReference(cellRef);
        Row row = sheet.getRow(cellReference.getRow());
        if( row == null ) {
            row = sheet.createRow(cellReference.getRow());
        }
        Cell cell = row.getCell(cellReference.getCol());
        if( cell == null ) {
            cell = row.createCell(cellReference.getCol());
        }
        return (SXSSFCell)cell;
    }


    private ExcelGenerator setStyleMMM(String name, Color foreColor, Color backColor, HorizontalAlignment ha, boolean bold) {
        XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.createCellStyle();
        Font font = workbook.createFont();
        XSSFFont xssfFont = (XSSFFont)font;
        if( foreColor != null ) {
            XSSFColor colorFore = new XSSFColor(foreColor, null);
            xssfFont.setColor(colorFore);
        }
        xssfFont.setBold(bold);

        cellStyle.setFont(font);

        if( backColor != null ) {
            XSSFColor colorBack = new XSSFColor(backColor, null);
            cellStyle.setFillForegroundColor(colorBack);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        cellStyle.setAlignment(ha);
        styles.put(name,cellStyle);

        //cellStyle.setBorderBottom(BorderStyle.THIN);
        return this;
    }

    public SXSSFWorkbook workbook;
    public ByteArrayOutputStream stream;
    public SXSSFSheet sheet ;
    public Map<String, XSSFCellStyle> styles = new HashMap<>();
    public Map<String, RecordSet> recordSets = new HashMap<>();
    public List<ReportParam> parameters = new LinkedList<>();

    public ExcelGenerator setPrameters(List<ReportParam> parameters) {
        this.parameters = parameters;
        return this;
    }

    public ExcelGenerator addParam(String name, String cell, String style, Object value) {
        ReportParam rp = new ReportParam();
        rp.setName(name).setCell(cell).setStyle(style).setValue(value);
        return this;
    }

    public ExcelGenerator prepareExcel() {
        stream = new ByteArrayOutputStream();
        workbook = new SXSSFWorkbook();
        sheet = workbook.createSheet();
        sheet.setDisplayGridlines(false);
        return this;
    }


    public ExcelGenerator finishExcel() throws Exception {
        workbook.write(stream);
        return this;
    }

    public ByteArrayOutputStream getStream() {
        return stream;
    }

    public ExcelGenerator saveToFile(String fileName) throws Exception {
        File targetFile = new File(fileName);
        FileOutputStream  fos = new FileOutputStream(targetFile);
        fos.write(stream.toByteArray());
        fos.close();
        return this;
    }

    public String resourceAsString( String recourceName ) {
        try {
            System.out.println( "Load resource:" + recourceName);
            InputStream isr = this.getClass().getClassLoader().getResourceAsStream(recourceName);
            int bufferSize = 1024;
            char[] buffer = new char[bufferSize];
            StringBuilder out = new StringBuilder();
            Reader in = new InputStreamReader(isr, StandardCharsets.UTF_8);
            for (int numRead; (numRead = in.read(buffer, 0, buffer.length)) > 0; ) {
                out.append(buffer, 0, numRead);
            }
            return out.toString();

        } catch( IOException ex ) {

            throw new UncheckedIOException("No existe el resource", ex);

        }

    }


}

