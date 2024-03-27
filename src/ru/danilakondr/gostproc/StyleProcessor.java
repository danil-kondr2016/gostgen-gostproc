package ru.danilakondr.gostproc;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.NumberingType;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.*;
import com.sun.star.container.XNameAccess;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

public class StyleProcessor extends Processor {
    private final XNameAccess xStyleFamilies;

    public StyleProcessor(XTextDocument xDoc) {
        super(xDoc);

        XStyleFamiliesSupplier xStyleSup = UnoRuntime.queryInterface(
                XStyleFamiliesSupplier.class,
                xDoc
        );
        xStyleFamilies = xStyleSup.getStyleFamilies();
    }

    @Override
    public void process() throws Exception {
        setPageStyles();
    }

    private void setPageStyles() throws Exception {
        Object oPageStyles = xStyleFamilies.getByName("PageStyles");
        XNameAccess xPageStyles = UnoRuntime.queryInterface(XNameAccess.class, oPageStyles);

        setPageStyle(xPageStyles, "Standard", true);
        setPageStyle(xPageStyles, "First Page", false);
    }

    private void setPageStyle(XNameAccess xPageStyles,
                              String styleName,
                              boolean footer
    )
    {
        try {
            XStyle xStyle = UnoRuntime.queryInterface(XStyle.class, xPageStyles.getByName(styleName));
            XPropertySet xStyleProp = UnoRuntime.queryInterface(XPropertySet.class, xStyle);

            xStyleProp.setPropertyValue("Size", new com.sun.star.awt.Size(21000, 29700));
            xStyleProp.setPropertyValue("LeftMargin", 3000);
            xStyleProp.setPropertyValue("RightMargin", 1500);
            xStyleProp.setPropertyValue("TopMargin", 2000);
            xStyleProp.setPropertyValue("BottomMargin", 2000);
            xStyleProp.setPropertyValue("FooterIsOn", footer);

            if (footer) {
                XText xFooterText = UnoRuntime
                        .queryInterface(XText.class,
                                xStyleProp.getPropertyValue("FooterText"));
                putPageNumber(xFooterText);
            }
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void putPageNumber(XText xFooterText) throws Exception {
        XTextCursor xCursor = xFooterText.createTextCursorByRange(xFooterText.getStart());

        xCursor.gotoStart(false);
        xCursor.gotoEnd(true);
        xFooterText.insertString(xCursor, "", true);

        xFooterText.insertString(xCursor, "\t", false);

        XTextField xPageNumber = createPageNumber();
        xFooterText.insertTextContent(xCursor, xPageNumber, false);
    }

    private XTextField createPageNumber() throws Exception {
        XMultiServiceFactory xMSF = UnoRuntime
                .queryInterface(XMultiServiceFactory.class, xDoc);

        Object oField = xMSF.createInstance("com.sun.star.text.textfield.PageNumber");
        XTextField xField = UnoRuntime.queryInterface(XTextField.class, oField);
        XPropertySet xFieldProp = UnoRuntime.queryInterface(XPropertySet.class, xField);

        xFieldProp.setPropertyValue("NumberingType", NumberingType.ARABIC);

        return xField;
    }
}
