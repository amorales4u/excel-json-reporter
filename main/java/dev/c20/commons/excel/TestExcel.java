package dev.c20.commons.excel;

import dev.c20.commons.excel.tools.JsonExcelReport;
import dev.c20.commons.excel.tools.RecordSet;
import dev.c20.commons.excel.tools.ReportParam;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

import java.io.*;
import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class TestExcel extends ExcelGeneratorFromJSON {


    static public void main(String[] args)  {

        try {
            Connection conn = getConnection();

            TestExcel atencionClientes = new TestExcel();
            atencionClientes.setConnection(conn);
            JsonExcelReport jsonExcelReport = new JsonExcelReport();

            jsonExcelReport.getParams().add( new ReportParam()
                    .setName("FechaDesde")
                    //.setCell("C2")
                    //.setStyle("title")
                    .setValue("2021/06/08")
            );
            jsonExcelReport.getParams().add( new ReportParam()
                    .setName("FechaHasta")
                    //.setCell("D2")
                    //.setStyle("title")
                    .setValue("2021/06/08")
            );

            /*
            jsonExcelReport.getParams().add( new ReportParam()
                    .setName("segmento")
                    //.setCell("E2")
                    //.setStyle("title")
                    .setValue("CBP")
            );

             */

            jsonExcelReport.setResourceJson("reports/test.json");

            atencionClientes.createReport(jsonExcelReport);

            ByteArrayOutputStream bos = atencionClientes.getStream();
            atencionClientes.saveToFile("target/test.xlsx");

            conn.close();
        } catch (Exception ex ) {
            ex.printStackTrace();
        }

    }

    static public Connection getConnection() {
        try {
            Class.forName("com.mysql.jdbc.Driver");
            Connection conn = DriverManager.getConnection(
                    "jdbc:mysql://localhost:3306/AMT_STG?serverTimezone=UTC","root","rn6jt2"
            );
            return conn;
        } catch( Exception ex ) {
            ex.printStackTrace();
        }

        return null;
    }


    @Override
    public void runCustomRecordSet( Map<String,Object>  recordSetCfg ) {
        /*
        if( ((String)recordSetCfg.get("name")).equalsIgnoreCase("crosstab") ) {
            customFill();
        }

         */

    }

    private void customFill() {
        Row row = sheet.getRow(3);

        int rowNum = 4;
        Cell cell = null;

        int colNum = 3;
        List<Segmento> segmentos = new LinkedList<>();
        while( ( cell = row.getCell(colNum)) != null ) {
            System.out.println(cell.getStringCellValue() );
            segmentos.add( new Segmento( colNum, cell.getStringCellValue()));
            colNum += 5;
        }

        System.out.println( "Totales en " + colNum );

        List<Test> rows = new LinkedList<>();
        String lastEntidad = null;
        Row lastEntidadRow = null;
        Row lastRow = null;
        while( (row = sheet.getRow(rowNum++) ) != null ) {
            lastRow = row;
            cell = row.getCell(1);
            if( cell == null ) {
                break;
            }
            System.out.println(cell.getStringCellValue() + " => " + cell.getCellStyle().getFillForegroundColor());
            if( cell.getCellStyle().getFillForegroundColor() == 0 ) {
                lastEntidad = cell.getStringCellValue();
                if( lastEntidadRow != null ) {

                    System.out.println("sum(" + (lastEntidadRow.getRowNum() + 1) + ":" + (row.getRowNum()-1));
                    for( int c = 3; c <= colNum + 4; c++ ) {
                        String columnLetter = CellReference.convertNumToColString(c);
                        Cell sumCell = lastEntidadRow.createCell(c);
                        sumCell.setCellStyle(styles.get("header-num"));
                        sumCell.setCellFormula("sum(" + columnLetter + (lastEntidadRow.getRowNum() + 2 ) + ":" + columnLetter + (row.getRowNum()) + ")");
                    }
                }
                lastEntidadRow = row;
                //rows.add( new Test(rowNum,lastEntidad,null) );
            } else {
                rows.add( new Test(row.getRowNum(),lastEntidad,cell.getStringCellValue()) );
            }

        }
        if( lastEntidadRow != null && lastRow != null) {
            System.out.println("sum(" + (lastEntidadRow.getRowNum() + 1) + ":" + (lastRow.getRowNum()));
            for( int c = 3; c <= colNum + 4; c++ ) {
                String columnLetter = CellReference.convertNumToColString(c);
                Cell sumCell = lastEntidadRow.createCell(c);
                sumCell.setCellStyle(styles.get("header-num"));
                sumCell.setCellFormula("sum(" + columnLetter + (lastEntidadRow.getRowNum() + 2) + ":" + columnLetter + (lastRow.getRowNum()+1) + ")");
            }
        }

        setCell(2,colNum,"Totales", "header-segmento-0");
        mergeRegion(colNum, 2, colNum + 4, 2);

        setCell(3,colNum,"Total", "header-num");
        setCell(3,colNum + 1,"En Tiempo", "header-en-tiempo");
        setCell(3,colNum + 2,"Vencidas", "header-en-vencida");
        setCell(3,colNum + 3,"No Atendidas", "header-no-atendida");
        setCell(3,colNum + 4,"Rechazadas", "header-rechazada");



        for( Test line : rows ) {
            System.out.print( line.rowNum + " " + line.entidad + " " + line.tipoSolicitud + " ");
            // para totales
            line.segmento.colNum = colNum;
            for( Segmento segmento : segmentos ) {
                setCell("A4",line.entidad);
                setCell("A5",line.tipoSolicitud);
                setCell("A6", segmento.segmento);
                RecordSet recordSet = sqlRunRecordSet("segmento-data",false,"row-solicitud-num");
                System.out.println( segmento.segmento + " start in col:" + segmento.colNum + " " + CellReference.convertNumToColString(segmento.colNum) + ":" + (line.rowNum) + " Record Count:" + recordSet.getRowCount());
                writeHorizontalRecordSet(line.rowNum,segmento.colNum,recordSet,5,false);
                if( recordSet.getRowCount() != 0) {
                    line.segmento.total += ((BigDecimal)recordSet.getRecords().get(0).get(0).getValue()).doubleValue();
                    line.segmento.enTiempo  += ((BigDecimal)recordSet.getRecords().get(0).get(1).getValue()).doubleValue();
                    line.segmento.vencidas  += ((BigDecimal)recordSet.getRecords().get(0).get(2).getValue()).doubleValue();
                    line.segmento.atendidas  += ((BigDecimal)recordSet.getRecords().get(0).get(3).getValue()).doubleValue();
                    line.segmento.rechazadas  += ((BigDecimal)recordSet.getRecords().get(0).get(4).getValue()).doubleValue();
                    line.segmento.eliminadas  += ((BigDecimal)recordSet.getRecords().get(0).get(5).getValue()).doubleValue();
                } else {
                    for( int x = segmento.colNum; x <= segmento.colNum+4; x ++ ) {
                        setCell(line.rowNum,x, "", "row-solicitud-num");
                    }

                }

            }
            System.out.println("");
        }


        for( Test line : rows ) {
            setCell( line.rowNum,line.segmento.colNum,((Object)Double.valueOf(line.segmento.total)),"row-solicitud-num");
            setCell( line.rowNum,line.segmento.colNum+1,((Object)Double.valueOf(line.segmento.enTiempo)),"row-solicitud-num");
            setCell( line.rowNum,line.segmento.colNum+2,((Object)Double.valueOf(line.segmento.vencidas)),"row-solicitud-num");
            setCell( line.rowNum,line.segmento.colNum+3,((Object)Double.valueOf(line.segmento.atendidas)),"row-solicitud-num");
            setCell( line.rowNum,line.segmento.colNum+4,((Object)Double.valueOf(line.segmento.rechazadas)),"row-solicitud-num");
        }


        setCell("A4","");
        setCell("A5","");
        setCell("A6", "");
        setCell(16,0,"");
    }

    private class Test {
        public Test( int rowNum, String entidad, String tipoSolicitud) {
            this.rowNum = rowNum;
            this.entidad = entidad;
            this.tipoSolicitud = tipoSolicitud;
            this.isEntidad = this.tipoSolicitud == null;
        }
        int rowNum;
        boolean isEntidad = false;
        String entidad;
        String tipoSolicitud;
        Segmento segmento = new Segmento();

    }

    private class Segmento {
        public Segmento() {

        }
        public Segmento(int colNum, String segmento) {
            this.colNum = colNum;
            this.segmento = segmento;
        }
        int colNum;
        String segmento;
        double total = 0;
        double enTiempo = 0;
        double vencidas = 0;
        double atendidas = 0;
        double rechazadas = 0;
        double eliminadas = 0;
    }


}
