package dev.c20.commons.excel.tools;


import java.util.LinkedList;
import java.util.List;

public class CrossTabTool {

    public List<List<String>> getCrossTabTwoLevels(List<String> firsts, List<String> seconds, List<String> columns) {
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
