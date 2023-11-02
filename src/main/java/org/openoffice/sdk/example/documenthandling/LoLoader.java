package org.openoffice.sdk.example.documenthandling;

import com.sun.star.beans.Pair;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.sheet.XSpreadsheetDocument;

import com.sun.star.uno.Exception;
import com.sun.star.uno.XComponentContext;
import de.hub.mse.office.CustomBootstrap;
import helper.*;

import java.util.*;

public class LoLoader {

    public static final int SHEET_WIDTH = 10;
    public static final int SHEET_DEPTH = 10;

    public static void main(String[] args) throws BootstrapException {

        String outFnm = null;
        if ((args.length < 1) || (args.length > 2)) {
            System.out.println("Usage: run ShowSheet fnm [out-fnm]");
            return;
        }
        if (args.length == 2)
            outFnm = args[1];

        XComponentContext context = CustomBootstrap.bootstrap(Bootstrap.getDefaultOptions());
        var loader = CustomBootstrap.getLoaderFromContext(context);

        XSpreadsheetDocument doc = Calc.openDoc(args[0], loader);
        if (doc == null) {
            System.out.println("Could not open " + args[0]);
            return;
        }

        GUI.setVisible(doc, true);
        Calc.gotoCell(doc, "A1");   // move view to top of sheet

        // XSpreadsheet sheet = Calc.getSheet(doc, 0);  // not needed here
        // var view = Calc.getView(doc);
        List<Pair<String, String>> errorCells = new ArrayList<>();
        Map<Integer, Integer> errorStatistics = new HashMap<>();
        for (var sheetName : doc.getSheets().getElementNames()) {
            System.out.println("Looking at sheet " + sheetName);
            var sheet = Calc.getSheet(doc, sheetName);
            Calc.setActiveSheet(doc, sheet);
            for (int rowNum = 0; rowNum < SHEET_DEPTH; rowNum++) {
                for (int columnNum = 0; columnNum < SHEET_WIDTH; columnNum++) {
                    var cell = Calc.getCell(sheet, columnNum, rowNum);
                    var val = Calc.getVal(cell);
                    var cellIdentifier = Calc.columnNumberStr(columnNum) + (rowNum + 1);
                    System.out.println("[" + cellIdentifier + "] " + (val == null ? "" : val.toString()));

                    if (cell.getError() != 0) {
                        errorCells.add(new Pair<>(sheetName, cellIdentifier + " -> " + cell.getError()));
                        var count = errorStatistics.getOrDefault(cell.getError(), 0);
                        errorStatistics.put(cell.getError(), ++count);
                    }
                }
            }
        }

        System.out.println("\n === Errorcells ===");
        for (var pair : errorCells) {
            System.out.println("[" + pair.First + "] " + pair.Second);
        }

        System.out.println("\n === Error statistics ===");
        var sortedSet = new TreeSet<Map.Entry<Integer, Integer>>(Comparator.comparingInt(Map.Entry::getKey));
        sortedSet.addAll(errorStatistics.entrySet());
        for (var entry : sortedSet) {
            System.out.println("Error " + entry.getKey() + ": " + entry.getValue());
        }
        Lo.waitEnter();  // wait for user to press <ENTER>
        if (outFnm != null)
            Lo.saveDoc(doc, outFnm);

        Lo.closeOffice();
    }

}
