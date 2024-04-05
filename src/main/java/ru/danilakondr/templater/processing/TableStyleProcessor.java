package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.*;
import com.sun.star.drawing.LineStyle;
import com.sun.star.table.BorderLine;
import com.sun.star.table.BorderLine2;
import com.sun.star.table.TableBorder;
import com.sun.star.table.TableBorder2;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextTablesSupplier;
import com.sun.star.uno.UnoRuntime;

import java.util.concurrent.atomic.AtomicInteger;

public class TableStyleProcessor extends Processor {
    public static final int TABLE_LINE_WIDTH = 17; // around 1/2 * (25.4/72) mm
    private final XNameAccess textTables;

    public TableStyleProcessor(XTextDocument xDoc) {
        super(xDoc);

        XTextTablesSupplier xSup = UnoRuntime
                .queryInterface(XTextTablesSupplier.class, xDoc);
        textTables = xSup.getTextTables();
    }

    @Override
    public void process() throws Exception {
        String[] tableNames = textTables.getElementNames();

        AtomicInteger i = new AtomicInteger(0);
        for (String tableName : tableNames) {
            System.out.printf("Processing table %d/%d...\n", i.incrementAndGet(), tableNames.length);
            XTextTable xTable = UnoRuntime
                    .queryInterface(XTextTable.class, textTables.getByName(tableName));
            processSingleTable(xTable);
        }
    }

    private void processSingleTable(XTextTable xTable) throws Exception {
        XPropertySet xTableProp = UnoRuntime
                .queryInterface(XPropertySet.class, xTable);
        TableBorder tableBorder = new TableBorder();

        tableBorder.HorizontalLine = new BorderLine();
        tableBorder.HorizontalLine.OuterLineWidth = TABLE_LINE_WIDTH;
        tableBorder.IsHorizontalLineValid = true;

        tableBorder.VerticalLine = new BorderLine();
        tableBorder.VerticalLine.OuterLineWidth = TABLE_LINE_WIDTH;
        tableBorder.IsVerticalLineValid = true;

        tableBorder.LeftLine = new BorderLine();
        tableBorder.LeftLine.OuterLineWidth = TABLE_LINE_WIDTH;
        tableBorder.IsLeftLineValid = true;

        tableBorder.RightLine = new BorderLine();
        tableBorder.RightLine.OuterLineWidth = TABLE_LINE_WIDTH;
        tableBorder.IsRightLineValid = true;

        tableBorder.TopLine = new BorderLine();
        tableBorder.TopLine.OuterLineWidth = TABLE_LINE_WIDTH;
        tableBorder.IsTopLineValid = true;

        tableBorder.BottomLine = new BorderLine();
        tableBorder.BottomLine.OuterLineWidth = TABLE_LINE_WIDTH;
        tableBorder.IsBottomLineValid = true;

        xTableProp.setPropertyValue("TableBorder", tableBorder);
    }
}
