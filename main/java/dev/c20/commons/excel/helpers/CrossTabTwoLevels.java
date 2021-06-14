package dev.c20.commons.excel.helpers;

import dev.c20.commons.excel.ExcelGeneratorFromJSON;
import dev.c20.commons.excel.tools.Field;
import dev.c20.commons.excel.tools.RecordSet;
import org.apache.poi.ss.util.CellReference;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class CrossTabTwoLevels {

    ExcelGeneratorFromJSON generator;
    Map<String,Object> config;
    int fromRow;
    int fromCol;
    int tabSize;
    String firstRecordSetName;
    String secondRecordSetName;
    String columnsRecordSetName;
    String crossDataRecordSetName;
    String crossDataRecordSetFirstLevelParamIn;
    String crossDataRecordSetSecondLevelParamIn;
    String crossDataRecordSetColumnLevelParamIn;

    String firstLevelStyle;
    String firstLevelStyleData;
    String secondLevelStyle;
    String dataStyle;
    List<String> mainTabStyle;
    List<String> tabStyles;
    boolean showVerticalName;
    boolean showVerticalDescription;
    String verticalLabel;

    boolean showHorizontalDescription;
    boolean showHorizontalName;
    String showTotalAt;
    String horizontalFormula;
    List<String> horizontalColumnLabels;

    public CrossTabTwoLevels(ExcelGeneratorFromJSON generator, Map<String,Object> config) {
        this.generator = generator;
        this.config = config;
        fromRow = (Integer)config.get("fromRow");
        fromCol = (Integer)config.get("fromCol");
        tabSize = (Integer)config.get("tabSize");

        Map<String,Object> recorSets = (Map<String, Object>) config.get("recordSets");

        firstRecordSetName = (String)recorSets.get("firstRecorSet");
        secondRecordSetName = (String)recorSets.get("secondRecordSet");
        columnsRecordSetName = (String)recorSets.get("columnsRecordSet");

        Map<String,String> crossDataRecordSet =  (Map<String,String>)recorSets.get("crossDataRecordSet");

        crossDataRecordSetName = crossDataRecordSet.get("recordSet");
        crossDataRecordSetFirstLevelParamIn = crossDataRecordSet.get("firstLevelParamIn");
        crossDataRecordSetSecondLevelParamIn = crossDataRecordSet.get("secondLevelParamIn");
        crossDataRecordSetColumnLevelParamIn= crossDataRecordSet.get("columnParamIn");

        Map<String,Object> styles =  (Map<String,Object>)config.get("styles");
        firstLevelStyle = (String)styles.get("firstLevelStyle");
        firstLevelStyleData = (String)styles.get("firstLevelStyleData");
        secondLevelStyle = (String)styles.get("secondLevelStyle");
        dataStyle = (String)styles.get("dataStyle");
        mainTabStyle = (List<String>)styles.get("mainTabStyle");
        tabStyles = (List<String>)styles.get("tabStyles");

        Map<String,Object> data = (Map<String, Object>) config.get("data");
        Map<String,Object> vertical = (Map<String, Object>) data.get("vertical");

        showVerticalName = (Boolean)vertical.get("showName");
        showVerticalDescription = (Boolean)vertical.get("showDescription");
        verticalLabel = (String)vertical.get("label");

        Map<String,Object> dataHorizontal = (Map<String, Object>) data.get("horizontal");
        showHorizontalDescription = (Boolean)dataHorizontal.get("showDescription");
        showHorizontalName = (Boolean)dataHorizontal.get("showName");
        showTotalAt = (String)dataHorizontal.get("showTotalAt");
        horizontalColumnLabels = (List<String>)dataHorizontal.get("labels");
        horizontalFormula = dataHorizontal.get("formula") + "(";
        readRecordsets();
        createCrossData();
        writeCrossDataToExcel();

    }

    RecordSet firstRecordSet;
    RecordSet secondRecordSet;
    RecordSet columnsRecordSet;
    RecordSet crossDataRecordSet;

    public CrossTabTwoLevels readRecordsets() {

        firstRecordSet = generator.sqlRunRecordSet(firstRecordSetName,false,null);
        secondRecordSet = generator.sqlRunRecordSet(secondRecordSetName,false,null);
        columnsRecordSet = generator.sqlRunRecordSet(columnsRecordSetName,false,null);
        //crossDataRecordSet = generator.sqlRunRecordSet(crossDataRecordSetName,false,null);

        return this;
    }

    public String getDeltaCellLetter( int col ) {
        return CellReference.convertNumToColString(col + fromCol);
    }

    public int getDeltaRowNum( int row ) {
        return row + fromRow;
    }

    public int getDeltaCol( int col ) {
        return col + fromCol;
    }

    public CrossCell setCell(CrossCellType type, Object value){
        return new CrossCell(type,value);
    }
    public CrossCell setCell(CrossCellType type, Object value, String style){
        return new CrossCell(type,value,style);
    }


    List<List<CrossCell>> crossData = new LinkedList<>();

    public CrossTabTwoLevels createCrossData() {

        List<List<CrossCell>> result = new LinkedList<>();

        if( showHorizontalDescription ) {
            result.add(createHorizontalHeaders());
        }

        result.add(createHorizontalTabs());

        int rowNum = result.size();

        for( List<Field> firstLevelRecord: firstRecordSet.getRecords() ) {
            List<Object> record = new LinkedList<>();
            result.add(getFirstLevelRecord(firstLevelRecord,rowNum ++));

            for( List<Field> secondLevelRecord : secondRecordSet.getRecords() ) {
                result.add( getSecondLevelRecord( firstLevelRecord, secondLevelRecord, rowNum ++));
            }
        }

        System.out.println("Cross form");
        for( List<CrossCell> row : result) {
            for( CrossCell col : row ) {
                System.out.print( col.value + "\t");
            }
            System.out.println("");
        }
        crossData = result;
        return this;
    }

    public List<CrossCell> createHorizontalHeaders() {
        List<CrossCell> record = new LinkedList<>();
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,null));
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,null) );
        }
        for( List<Field> columnsRecord : columnsRecordSet.getRecords() ) {
            record.add(setCell(CrossCellType.VALUE,columnsRecord.get(1).getValue()));
            for( int i = 1; i < tabSize; i ++ ) {
                record.add(setCell(CrossCellType.VALUE,null));
            }
        }
        record.add(setCell(CrossCellType.VALUE,"Total"));

        for( int i = 1; i < tabSize; i ++ ) {
            record.add(setCell(CrossCellType.VALUE,null));
        }

        return record;
    }

    public List<CrossCell> createHorizontalTabs() {
        List<CrossCell> record = new LinkedList<>();
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,"", firstLevelStyle) );
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,verticalLabel, firstLevelStyle) );
        }
        for( List<Field> columnsRecord : columnsRecordSet.getRecords() ) {
            if( showTotalAt.equalsIgnoreCase("start") ) {
                if( showHorizontalName ) {
                    record.add(setCell(CrossCellType.VALUE,columnsRecord.get(0).getValue(),firstLevelStyle));
                } else {
                    record.add(setCell(CrossCellType.VALUE,"Total",firstLevelStyle));
                }
            }
            for( int idx = 0; idx < horizontalColumnLabels.size(); idx ++ ) {
                String label = horizontalColumnLabels.get(idx);
                record.add(setCell(CrossCellType.VALUE,label,tabStyles.get(idx)));
            }
            if( showTotalAt.equalsIgnoreCase("finish") ) {
                if( showHorizontalName ) {
                    record.add(setCell( CrossCellType.VALUE,columnsRecord.get(0).getValue(),firstLevelStyle));
                } else {
                    record.add(setCell(CrossCellType.VALUE, "Total",firstLevelStyle));
                }
            }

        }

        if( showTotalAt.equalsIgnoreCase("start") ) {
            record.add(setCell(CrossCellType.VALUE,"Total", tabStyles.get(0)));
        }
        for( int idx = 0; idx < horizontalColumnLabels.size(); idx ++ ) {
            String label = horizontalColumnLabels.get(idx);
            record.add(setCell(CrossCellType.VALUE,label,tabStyles.get(idx)));
        }
        if( showTotalAt.equalsIgnoreCase("finish") ) {
            record.add(setCell(CrossCellType.VALUE,"Total", tabStyles.get(tabStyles.size()-1)));
        }
        return record;
    }

    public List<CrossCell> getFirstLevelRecord(List<Field> firstLevelRecord, int rowNum) {
        List<CrossCell> record = new LinkedList<>();
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,firstLevelRecord.get(0).getValue(), firstLevelStyle));
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,firstLevelRecord.get(1).getValue(),firstLevelStyle));
        }
        // Tab columns
        for( List<Field> columnRecord : columnsRecordSet.getRecords() ) {
            String cellLetter = getDeltaCellLetter( record.size() );
            if( showTotalAt.equalsIgnoreCase("start") ) {
                record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetter + ( getDeltaRowNum( rowNum + 2 ) ) +
                                ":" + cellLetter + getDeltaRowNum( rowNum + secondRecordSet.getRowCount() + 1) + ")"
                        , firstLevelStyleData));
            }
            for( int c = 0; c < horizontalColumnLabels.size(); c ++ ) {
                String cl = getDeltaCellLetter( record.size() );
                record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cl + getDeltaRowNum( rowNum + 2 ) +
                                ":" + cl + getDeltaRowNum( rowNum + secondRecordSet.getRowCount() + 1) + ")",
                        firstLevelStyleData));
            }
            if( showTotalAt.equalsIgnoreCase("finish") ) {
                record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetter + getDeltaRowNum( rowNum + 2 ) +
                                ":" + cellLetter + getDeltaRowNum( rowNum + secondRecordSet.getRowCount() + 1) + ")",
                        firstLevelStyleData));
            }

        }
        // totales
        String cellLetter = getDeltaCellLetter( record.size() );
        if( showTotalAt.equalsIgnoreCase("start") ) {
            record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetter + getDeltaRowNum(rowNum + 2) +
                            ":" + cellLetter + getDeltaRowNum(rowNum + secondRecordSet.getRowCount() + 1) + ")"
                    , firstLevelStyleData));
        }
        for( int c = 0; c < horizontalColumnLabels.size(); c ++ ) {
            String cl = getDeltaCellLetter( record.size()  );
            record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cl + getDeltaRowNum( rowNum + 2 ) +
                            ":" + cl + getDeltaRowNum( rowNum + secondRecordSet.getRowCount() + 1) + ")",
                    firstLevelStyleData));
        }
        if( showTotalAt.equalsIgnoreCase("finish") ) {
            record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetter + getDeltaRowNum(rowNum + 2) +
                            ":" + cellLetter + getDeltaRowNum(rowNum + secondRecordSet.getRowCount() + 1) + ")",
                    firstLevelStyleData));
        }

        return record;
    }


    public List<CrossCell> getSecondLevelRecord(List<Field> firstLevelRecord, List<Field> secondLevelRecord,int rowNum) {
        List<CrossCell> record = new LinkedList<>();
        List<List<String>> granTotal = new LinkedList<>();
        for( int c = 0; c < horizontalColumnLabels.size(); c ++ ) {
            List<String> columnsForTotal = new LinkedList<>();
            granTotal.add(columnsForTotal);
        }
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,secondLevelRecord.get(0).getValue(),secondLevelStyle));
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,secondLevelRecord.get(1).getValue(),secondLevelStyle));
        }
        // Tab columns

        int idxTotal = 0;
        for( List<Field> columnRecord : columnsRecordSet.getRecords() ) {
            if( showTotalAt.equalsIgnoreCase("start") ) {
                String cellLetterStart = getDeltaCellLetter(record.size() + 1);
                String cellLetterFinish = getDeltaCellLetter(record.size() + horizontalColumnLabels.size());
                record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetterStart + getDeltaRowNum(rowNum + 1) +
                                ":" + cellLetterFinish + getDeltaRowNum(rowNum + 1) + ")",
                        dataStyle));
            }
            for( int c = 0; c < horizontalColumnLabels.size(); c ++ ) {
                List<String> totales = granTotal.get(c);
                totales.add(getDeltaCellLetter(record.size() + c ));
            }

            for( int c = 0; c < horizontalColumnLabels.size(); c ++ ) {
                if( c == 0 ) {
                    CrossCell cell = setCell(CrossCellType.READ,dataStyle);
                    cell.params.add(firstLevelRecord.get(0).getValue());
                    cell.params.add(secondLevelRecord.get(0).getValue());
                    cell.params.add(columnRecord.get(0).getValue());
                    record.add(cell);
                } else {
                    CrossCell cell = setCell(CrossCellType.EMPTY,dataStyle);
                    record.add(cell);
                }
            }

            if( showTotalAt.equalsIgnoreCase("finish") ) {
                String cellLetterStart = getDeltaCellLetter(record.size() - horizontalColumnLabels.size() );
                String cellLetterFinish = getDeltaCellLetter(record.size() - 1 );
                record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetterStart + getDeltaRowNum(rowNum + 1) +
                                ":" + cellLetterFinish + getDeltaRowNum(rowNum + 1) + ")",
                        dataStyle));
            }

            idxTotal++;

        }
        // totales
        int colIdx = 0;
        if( showTotalAt.equalsIgnoreCase("start") ) {
            String cellLetterStart = getDeltaCellLetter(record.size() + 1 );
            String cellLetterFinish = getDeltaCellLetter(record.size() + horizontalColumnLabels.size() );
            record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetterStart + getDeltaRowNum(rowNum + 1) +
                            ":" + cellLetterFinish + getDeltaRowNum(rowNum + 1) + ")",
                    dataStyle));
            colIdx = 1;
        }
        int fromColIdx = fromCol;
        if( showVerticalName ) {
            fromColIdx++;
        }

        if( showVerticalDescription ) {
            fromColIdx ++;
        }
        int horizontalTabsCount = horizontalColumnLabels.size() + 1;
        for( List<String> totalForRow : granTotal ) {
            String cellFormula = "sum(" + listAsString(totalForRow,getDeltaRowNum( rowNum + 1) +",") + ")";
            record.add(setCell(CrossCellType.FORMULA,cellFormula,dataStyle));

        }
        /*
        for( int c = 1; c <= horizontalTabsCount; c ++ ) {
            String cellFormula = "sum(";
            String rowStr = getDeltaRowNum( rowNum + 1) + "" ;
            for( int column = 1; column <= columnsRecordSet.getRowCount(); column++  ) {
                int colCalc = ( fromColIdx + (horizontalTabsCount * (column  * c )) ) + colIdx - 1;
                cellFormula += CellReference.convertNumToColString( colCalc ) + rowStr;
                if(column == columnsRecordSet.getRowCount() - 1) {
                    cellFormula +=")";
                } else {
                    cellFormula +=",";
                }
            }
            record.add(setCell(CrossCellType.FORMULA,cellFormula, dataStyle));
        }

         */

        if( showTotalAt.equalsIgnoreCase("finish") ) {
            String cellLetterStart = getDeltaCellLetter(record.size() - horizontalColumnLabels.size());
            String cellLetterFinish = getDeltaCellLetter(record.size() -1 );
            record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetterStart + getDeltaRowNum(rowNum + 1) +
                            ":" + cellLetterFinish + getDeltaRowNum(rowNum + 1) + ")",
                    dataStyle));
        }

        return record;
    }

    static public String listAsString( List<String > list, String separator ) {
        String result = "";

        for( String line : list ) {
            result += line + separator;
        }

        return result.substring(0,result.length()-(separator.length()-1));
    }

    public CrossTabTwoLevels writeCrossDataToExcel() {
        int indexRow = 0;
        if( showHorizontalDescription ) {
            List<CrossCell> column = crossData.get(0);

            int styleIdx = 0;
            for( int colIdx = 0; colIdx < column.size(); colIdx ++ ) {
                CrossCell cell = column.get(colIdx);
                if( cell.value != null ) {
                    if( styleIdx == 0 ) {
                        styleIdx = 1;
                    } else {
                        styleIdx = 0;
                    }
                    this.generator.mergeRegion(getDeltaCol(colIdx),getDeltaRowNum(0)
                            ,getDeltaCol(colIdx + horizontalColumnLabels.size()), getDeltaRowNum(0));
                    this.generator.setCell( getDeltaRowNum(0), getDeltaCol(colIdx), cell.value, mainTabStyle.get(styleIdx));
                }
            }
            indexRow++;
        }

        for( int rowIdx = indexRow; rowIdx < crossData.size(); rowIdx ++ ) {
            List<CrossCell> columns = crossData.get(rowIdx);
            for( int colIdx = 0; colIdx < columns.size(); colIdx ++ ) {
                CrossCell column = columns.get(colIdx);
                switch( column.type ) {
                    case EMPTY:
                        //this.generator.setCellStyle( getDeltaRowNum(rowIdx), getDeltaCol(colIdx),  column.style );
                        break;
                    case VALUE:
                        this.generator.setCell( getDeltaRowNum(rowIdx), getDeltaCol(colIdx), column.value, column.style );
                        break;
                    case FORMULA:
                        this.generator.setFormula(getDeltaRowNum(rowIdx), getDeltaCol(colIdx), (String)column.value, column.style );
                        break;
                    case READ:

                        // first level param
                        this.generator.setCell(crossDataRecordSetFirstLevelParamIn,column.params.get(0));
                        // second level param
                        this.generator.setCell(crossDataRecordSetSecondLevelParamIn,column.params.get(1));
                        // column level param
                        this.generator.setCell(crossDataRecordSetColumnLevelParamIn,column.params.get(2));
                        RecordSet rs = this.generator.sqlRunRecordSet(crossDataRecordSetName,false,dataStyle);
                        if( rs.getRowCount() == 1 ) {
                            this.generator.writeVerticalRecordSet(getDeltaRowNum(rowIdx), getDeltaCol(colIdx), rs);
                        } else {
                            for( int col = 0; col < horizontalColumnLabels.size(); col ++  ) {
                                this.generator.setCellStyle(getDeltaRowNum(rowIdx), getDeltaCol(col+colIdx) , dataStyle);
                            }
                        }

                        this.generator.setCell(crossDataRecordSetFirstLevelParamIn,"");
                        // second level param
                        this.generator.setCell(crossDataRecordSetSecondLevelParamIn,"");
                        // column level param
                        this.generator.setCell(crossDataRecordSetColumnLevelParamIn,"");

                        break;
                }
            }

        }


        return this;
    }

    static public List<List<String>> getCrossTabTwoLevels(List<String> firsts, List<String> seconds, List<String> columns) {
        // sample
        // BN   null AI BEI BP  BPV
        // BN   AD  AI BEI BP  BPV
        // BN   AO  AI BEI BP  BPV
        // CB   null AI BEI BP  BPV
        // CB   AD  AI BEI BP  BPV
        // CB   AO  AI BEI BP  BPV
        List<List<String>> result = new LinkedList<>();
        List<String> cols = new LinkedList<>();
        cols.add(null);
        cols.add(null);

        for( String col : columns) {
            cols.add(col);
        }

        result.add(cols);

        for( String first : firsts ) {
            List<String> row = new LinkedList<>();
            row.add(first);
            row.add(null);
            for( String column : columns ) {
                row.add(" X ");
            }
            result.add(row);
            for( String second : seconds ) {
                List<String> sec = new LinkedList<>();
                sec.add(first);
                sec.add(second);
                for( String column : columns ) {
                    sec.add(" X ");
                }
                result.add(sec);
            }
        }

        for( List<String> row : result) {
            for( String col : row ) {
                System.out.print( col + "\t");
            }
            System.out.println("");
        }

        return result;

    }
}
