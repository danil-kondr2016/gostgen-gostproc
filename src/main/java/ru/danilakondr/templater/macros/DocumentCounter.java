package ru.danilakondr.templater.macros;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XController;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextSectionsSupplier;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import org.apache.commons.text.lookup.StringLookup;


public class DocumentCounter implements StringLookup {
    private final XTextDocument xDoc;

    private final int n_pages;

    private final int n_figures;

    private final int n_tables;

    public DocumentCounter(XTextDocument xDoc) {
        this.xDoc = xDoc;
        this.n_figures = getFigureCount();
        this.n_pages = getPageCount();
        this.n_tables = getTableCount();
    }

    private int getPageCount() {
        try {
            XController xController = xDoc.getCurrentController();
            XPropertySet xCtrlProp = UnoRuntime
                    .queryInterface(XPropertySet.class, xController);

            return AnyConverter
                    .toInt(xCtrlProp.getPropertyValue("PageCount"));
        }
        catch (Exception e) {
            return 0;
        }
    }

    private int getFigureCount() {
        try {
            XEnumerationAccess xEnumAccess = UnoRuntime
                    .queryInterface(XEnumerationAccess.class, xDoc.getText());
            XEnumeration xEnum = xEnumAccess.createEnumeration();

            int count = 0;
            while (xEnum.hasMoreElements()) {
                XTextContent xParagraph = UnoRuntime
                        .queryInterface(XTextContent.class, xEnum.nextElement());
                XPropertySet xParProp = UnoRuntime
                        .queryInterface(XPropertySet.class, xParagraph);

                String styleName = AnyConverter
                        .toString(xParProp.getPropertyValue("ParaStyleName"));
                if (styleName.equals("FigureWithCaption"))
                    count++;
            }
            return count;
        }
        catch (Exception e) {
            return 0;
        }
    }

    private int getTableCount() {
        int count = 0;
        XTextSectionsSupplier xSup = UnoRuntime
                .queryInterface(XTextSectionsSupplier.class, xDoc);
        XNameAccess xSections = xSup.getTextSections();

        for (String objId : xSections.getElementNames()) {
            if (objId.startsWith("tbl:"))
                count++;
        }

        return count;
    }

    @Override
    public String lookup(String s) {
        if (s.compareTo("N_PAGES") == 0)
            return String.valueOf(n_pages);
        if (s.compareTo("N_FIGURES") == 0)
            return String.valueOf(n_figures);
        if (s.compareTo("N_TABLES") == 0)
            return String.valueOf(n_tables);

        return null;
    }
}
