package dev.c20.commons.excel.helpers;

import dev.c20.commons.excel.ExcelGeneratorFromJSON;
import dev.c20.commons.excel.tools.Field;
import dev.c20.commons.excel.tools.RecordSet;
import org.apache.poi.ss.util.CellReference;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class CrossTabOneLevel {

    ExcelGeneratorFromJSON generator;
    Map<String,Object> config;

    RecordSet firstRecordSet;
    RecordSet secondRecordSet;
    RecordSet columnsRecordSet;
    RecordSet crossDataRecordSet;

    int fromRow;
    int fromCol;
    int tabSize;

    String firstRecordSetName;
    String secondRecordSetName;
    String columnsRecordSetName;

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


    String crossDataRecordSetName;
    String crossDataRecordSetFirstLevelParamIn;
    String crossDataRecordSetSecondLevelParamIn;
    String crossDataRecordSetColumnLevelParamIn;

    List<List<CrossCell>> crossData = new LinkedList<>();

    public CrossTabOneLevel(ExcelGeneratorFromJSON generator, Map<String,Object> config) {
        this.generator = generator;
        this.config = config;

        fromRow = (Integer)config.get("fromRow");
        fromCol = (Integer)config.get("fromCol");
        tabSize = (Integer)config.get("tabSize");

        Map<String,Object> recorSets = (Map<String, Object>) config.get("recordSets");

        firstRecordSetName = (String)recorSets.get("firstRecorSet");
        secondRecordSetName = (String)recorSets.get("secondRecordSet");
        columnsRecordSetName = (String)recorSets.get("columnsRecordSet");

        Map<String,Object> styles =  (Map<String,Object>)config.get("styles");
        firstLevelStyle = (String)styles.get("firstLevelStyle");
        firstLevelStyleData = (String)styles.get("firstLevelStyleData");
        secondLevelStyle = (String)styles.get("secondLevelStyle");
        dataStyle = (String)styles.get("dataStyle");
        mainTabStyle = (List<String>)styles.get("mainTabStyle");
        tabStyles = (List<String>)styles.get("tabStyles");

        Map<String,String> crossDataRecordSet =  (Map<String,String>)recorSets.get("crossDataRecordSet");

        crossDataRecordSetName = crossDataRecordSet.get("recordSet");
        crossDataRecordSetFirstLevelParamIn = crossDataRecordSet.get("firstLevelParamIn");
        crossDataRecordSetSecondLevelParamIn = crossDataRecordSet.get("secondLevelParamIn");
        crossDataRecordSetColumnLevelParamIn= crossDataRecordSet.get("columnParamIn");

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

    public CrossTabOneLevel readRecordsets() {

        firstRecordSet = generator.sqlRunRecordSet(firstRecordSetName,false,null);
        secondRecordSet = generator.sqlRunRecordSet(secondRecordSetName,false,null);
        columnsRecordSet = generator.sqlRunRecordSet(columnsRecordSetName,false,null);
        crossDataRecordSet = generator.sqlRunRecordSet(crossDataRecordSetName,false,null);

        return this;
    }

    public CrossTabOneLevel createCrossData() {

        List<List<CrossCell>> result = new LinkedList<>();

        if (showHorizontalDescription) {
            result.add(createHorizontalHeaders());
        }

        result.add(createHorizontalTabs());

        int rowNum = result.size();

        for (List<Field> secondLevelRecord : secondRecordSet.getRecords()) {
            result.add(getSecondLevelRecord( secondLevelRecord, rowNum++));
        }

        System.out.println("Cross form");
        for (List<CrossCell> row : result) {
            for (CrossCell col : row) {
                System.out.print(col.type + "-" + col.style +"-" + col.value + "\t");
            }
            System.out.println("");
        }
        crossData = result;
        return this;
    }


    public List<CrossCell> getSecondLevelRecord( List<Field> secondLevelRecord,int rowNum) {
        List<CrossCell> record = new LinkedList<>();
        List<String> granTotal = new LinkedList<>();
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,secondLevelRecord.get(0).getValue(),secondLevelStyle));
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,secondLevelRecord.get(1).getValue(),secondLevelStyle));
        }

        // total

        CrossCell superTotal = new CrossCell();
        superTotal.value = horizontalFormula;
        superTotal.style = dataStyle;
        int superTotalRow = record.size();

        superTotal.type = CrossCellType.FORMULA;
        if( showTotalAt.equalsIgnoreCase("start") ) {
            String cellLetterStart = getDeltaCellLetter(record.size() + 1 + 1);
            String cellLetterFinish = getDeltaCellLetter(record.size() + columnsRecordSet.getRowCount() +1);


            record.add(superTotal);
        }
        // Tab columns

        int idxTotal = 0;
        for( List<Field> firstLevelRecord : firstRecordSet.getRecords() ) {
            // total de first level
            if( showTotalAt.equalsIgnoreCase("start") ) {
                String cellLetterStart = getDeltaCellLetter(record.size() + 1  );
                String cellLetterFinish = getDeltaCellLetter(record.size() + columnsRecordSet.getRowCount() );

                superTotal.value += getDeltaCellLetter(record.size() ) + getDeltaRowNum(rowNum + 1) + ",";

                record.add(setCell(CrossCellType.FORMULA,horizontalFormula + cellLetterStart + getDeltaRowNum(rowNum + 1) +
                                ":" + cellLetterFinish + getDeltaRowNum(rowNum + 1) + ")",
                        dataStyle));
            }

            for( int c = 0; c < columnsRecordSet.getRowCount(); c ++ ) {
                Field columnRecord = columnsRecordSet.getRecords().get(c).get(0);
                CrossCell cell = setCell(CrossCellType.VALUE,null,dataStyle);
                cell.value = findInCrossData( crossDataRecordSet.getRecords(), firstLevelRecord.get(0).getValue(), secondLevelRecord.get(0).getValue(), columnRecord.getValue());
                record.add(cell);
            }

            idxTotal++;

        }
        String superTotalFormula = (String)superTotal.value;
        superTotal.value = superTotalFormula.substring(0,superTotalFormula.length()-1) + " )";
        record.set(superTotalRow, superTotal);
        return record;
    }

    // utils

    public Object findInCrossData( List<List<Field>> crossData, Object firstLevelRecord, Object secondLevelRecord, Object columnRecord) {
        Object result = null;

        for( List<Field> record : crossData ) {
            if( record.get(0).getValue().equals(firstLevelRecord) &&
                    record.get(1).getValue().equals(secondLevelRecord) &&
                    record.get(2).getValue().equals(columnRecord) ) {
                return record.get(3).getValue();
            }
        }

        return result;
    }
    public CrossTabOneLevel writeCrossDataToExcel() {
        int indexRow = 0;

        if( showHorizontalDescription ) {
            List<CrossCell> column = crossData.get(0);

            int styleIdx = 0;

            for( int colIdx = 3; colIdx < column.size(); colIdx ++ ) {
                CrossCell cell = column.get(colIdx);
                if( cell.value != null ) {
                    if( styleIdx == 0 ) {
                        styleIdx = 1;
                    } else {
                        styleIdx = 0;
                    }
                    this.generator.mergeRegion(getDeltaCol(colIdx),getDeltaRowNum(0)
                            ,getDeltaCol(colIdx + columnsRecordSet.getRecords().size()), getDeltaRowNum(0));

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
                            for( int col = 0; col < columnsRecordSet.getRecords().size(); col ++  ) {
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

    public List<CrossCell> createHorizontalHeaders() {
        List<CrossCell> record = new LinkedList<>();
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,null));
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,null) );
        }

        record.add(setCell(CrossCellType.VALUE,"Total"));

        for( List<Field> columnsRecord : firstRecordSet.getRecords() ) {
            record.add(setCell(CrossCellType.VALUE,columnsRecord.get(1).getValue()));
            for( int i = 1; i <= columnsRecordSet.getRowCount(); i ++ ) {
                record.add(setCell(CrossCellType.VALUE,null));
            }
        }

        return record;
    }
    public List<CrossCell> createHorizontalTabs() {
        List<CrossCell> record = new LinkedList<>();
        if( showVerticalName ) {
            record.add( setCell(CrossCellType.VALUE,verticalLabel, firstLevelStyle) );
        }
        if( showVerticalDescription ) {
            record.add( setCell(CrossCellType.VALUE,"", firstLevelStyle) );
        }
        if( showTotalAt.equalsIgnoreCase("start") ) {
            record.add(setCell(CrossCellType.VALUE,"Total",firstLevelStyle));
        }
        for( List<Field> firstLevelRecord : firstRecordSet.getRecords() ) {
            if( showTotalAt.equalsIgnoreCase("start") ) {
                if( showHorizontalName ) {
                    record.add(setCell(CrossCellType.VALUE,firstLevelRecord.get(1).getValue(),firstLevelStyle));
                } else {
                    record.add(setCell(CrossCellType.VALUE,firstLevelRecord.get(0).getValue(),firstLevelStyle));
                }
            }
            for( int idx = 0; idx < columnsRecordSet.getRecords().size(); idx ++ ) {
                if( showHorizontalName ) {
                    record.add(setCell(CrossCellType.VALUE, columnsRecordSet.get(idx, 1).getValue(), secondLevelStyle));
                } else {
                    record.add(setCell(CrossCellType.VALUE, columnsRecordSet.get(idx, 0).getValue(), secondLevelStyle));
                }
            }
            if( showTotalAt.equalsIgnoreCase("finish") ) {
                if( showHorizontalName ) {
                    record.add(setCell( CrossCellType.VALUE,firstLevelRecord.get(1).getValue(),firstLevelStyle));
                } else {
                    record.add(setCell(CrossCellType.VALUE, "Total",firstLevelStyle));
                }
            }

        }

        if( showTotalAt.equalsIgnoreCase("finish") ) {
            record.add(setCell(CrossCellType.VALUE,"Total", tabStyles.get(tabStyles.size()-1)));
        }
        return record;
    }

}
