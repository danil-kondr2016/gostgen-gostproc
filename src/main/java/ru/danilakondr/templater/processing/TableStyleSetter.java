package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.table.BorderLine;
import com.sun.star.table.TableBorder;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.UnoRuntime;

import java.util.function.Consumer;

public class TableStyleSetter implements TextDocument.ObjectProcessor<XTextTable> {
    public static final int TABLE_LINE_WIDTH = 17; // around 1/2 * (25.4/72) mm
    @Override
    public void process(XTextTable xTable, XTextDocument xDoc) {
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

        try {
            xTableProp.setPropertyValue("TableBorder", tableBorder);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
