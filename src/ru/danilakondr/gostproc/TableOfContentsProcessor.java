package ru.danilakondr.gostproc;

import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

public class TableOfContentsProcessor implements Processor {
    private final XTextDocument xDoc;
    private final XText xText;
    private final XParagraphCursor xCursor;

    public TableOfContentsProcessor(XTextDocument xDoc) {
        this.xDoc = xDoc;
        this.xText = xDoc.getText();
        XTextCursor xTextCursor = this.xText.createTextCursorByRange(this.xText.getStart());
        this.xCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
    }

    @Override
    public void process() throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("^([ \\t]*)%TOC%$");
        xSD.setPropertyValue("SearchRegularExpression", true);

        Object oResult = xS.findFirst(xSD);
        if (oResult == null)
            return;

        XTextRange xRange = UnoRuntime.queryInterface(XTextRange.class, oResult);
        xCursor.gotoRange(xRange.getStart(), false);
        xCursor.gotoRange(xRange.getEnd(), true);
        putTableOfContents(xCursor);
    }

    private Object createIndex() throws Exception {
        XMultiServiceFactory xMSF = UnoRuntime
                .queryInterface(XMultiServiceFactory.class, xDoc);

        return xMSF.createInstance("com.sun.star.text.ContentIndex");
    }

    private void putTableOfContents(XTextCursor cursor) throws Exception {
        Object oIndex = createIndex();

        XDocumentIndex xIndex = UnoRuntime
                .queryInterface(XDocumentIndex.class, oIndex);
        XPropertySet xIndexProp = UnoRuntime
                .queryInterface(XPropertySet.class, xIndex);

        XText xText = xDoc.getText();

        xText.insertTextContent(cursor, xIndex, true);

        // Сначала добавить и только потом выставлять свойства!
        xIndexProp.setPropertyValue("CreateFromOutline", true);
        xIndexProp.setPropertyValue("Title", "Оглавление");

        xIndex.update();
    }
}
