package ru.danilakondr.md2writer;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;

import static com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK;


//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws java.lang.Exception {
        XComponentContext xContext = Bootstrap.bootstrap();

        XMultiComponentFactory xMCF = xContext.getServiceManager();

        Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext);

        XDesktop xDesktop = (XDesktop) UnoRuntime.queryInterface(
                XDesktop.class, oDesktop);

        XComponentLoader xCompLoader = (XComponentLoader) UnoRuntime
                .queryInterface(com.sun.star.frame.XComponentLoader.class, xDesktop);

        PropertyValue[] val = new PropertyValue[0];
        XComponent xComp = xCompLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, val);

        XTextDocument xDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);
        XText xText = xDoc.getText();
        XSimpleText xSText = UnoRuntime.queryInterface(XSimpleText.class, xText);
        XTextCursor xCursor = xText.createTextCursorByRange(xText.getStart());
        XParagraphCursor xParCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xCursor);

        xSText.insertString(xCursor, "First pangram", false);
        xSText.insertControlCharacter(xCursor, PARAGRAPH_BREAK, false);
        xSText.insertString(xCursor, "The quick brown fox jumps over the lazy dog.", false);
        xSText.insertControlCharacter(xCursor, PARAGRAPH_BREAK, false);
        xSText.insertString(xCursor, "Second pangram", false);
        xSText.insertControlCharacter(xCursor, PARAGRAPH_BREAK, false);
        xSText.insertString(xCursor, "Jackdaws love my big sphinx of quartz", false);

        xParCursor.gotoStart(false);
        xParCursor.gotoEndOfParagraph(true);

        XPropertySet xParProp = UnoRuntime.queryInterface(XPropertySet.class, xParCursor);
        xParProp.setPropertyValue("ParaStyleName", "Heading 1");

        xParCursor.gotoNextParagraph(false);
        xParCursor.gotoEndOfParagraph(true);
        xParProp.setPropertyValue("ParaStyleName", "Text body");

        xParCursor.gotoNextParagraph(false);
        xParCursor.gotoEndOfParagraph(true);
        xParProp.setPropertyValue("ParaStyleName", "Heading 1");

        xParCursor.gotoNextParagraph(false);
        xParCursor.gotoEndOfParagraph(true);
        xParProp.setPropertyValue("ParaStyleName", "Text body");

        XMultiServiceFactory xDocMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, xDoc);
        Object oIndex = xDocMSF.createInstance("com.sun.star.text.ContentIndex");
        XDocumentIndex xIndex = UnoRuntime.queryInterface(XDocumentIndex.class, oIndex);
        XPropertySet xIndexProp = UnoRuntime.queryInterface(XPropertySet.class, oIndex);

        xParCursor.gotoEnd(false);
        xText.insertTextContent(xParCursor, xIndex, false);
        xIndexProp.setPropertyValue("CreateFromOutline", true);
        xIndexProp.setPropertyValue("Title", "Оглавление");
        xIndex.update();

        xDesktop.terminate();
    }
}