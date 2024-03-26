package ru.danilakondr.md2writer;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

public class TableOfContentsProcessor {
    private final XTextDocument xDoc;

    public TableOfContentsProcessor(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    public void process() throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setPropertyValue("SearchString", "^([ \\t]*)%TOC%$");
        xSD.setPropertyValue("SearchCaseSensitive", true);
        xSD.setPropertyValue("SearchWords", true);
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess xResults = xS.findAll(xSD);
        XText xText = xDoc.getText();
        XSimpleText xSText = UnoRuntime
                .queryInterface(XSimpleText.class, xText);
        for (int i = 0; i < xResults.getCount(); i++) {
            XTextCursor cur = UnoRuntime
                    .queryInterface(XTextCursor.class, xResults.getByIndex(i));
            putTableOfContents(cur);
        }
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

        // TODO implement removing %TOC% in the cursor position
        xText.insertTextContent(cursor, xIndex, true);

        // Сначала добавить и только потом выставлять свойства!
        xIndexProp.setPropertyValue("CreateFromOutline", true);
        xIndexProp.setPropertyValue("Title", "Оглавление");
        xIndex.update();
    }
}
