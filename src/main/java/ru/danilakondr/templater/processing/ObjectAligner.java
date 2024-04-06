package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;

public class ObjectAligner implements TextDocument.ObjectProcessor<Object> {
    @Override
    public void process(Object object, XTextDocument xDoc) {
        try {
            XTextContent xContent = UnoRuntime
                    .queryInterface(XTextContent.class, object);
            XTextRange xRange = xContent.getAnchor();

            if (xRange.getText() != xDoc.getText())
                return;

            XTextCursor xCursor = xRange.getText().createTextCursorByRange(xRange);
            if (hasTextAroundFormula(xRange, xDoc))
                return;

            xCursor.gotoRange(xRange, true);

            XPropertySet xCursorProps = UnoRuntime
                    .queryInterface(XPropertySet.class, xCursor);

            try {
                String numStyle = (String)xCursorProps.getPropertyValue("NumberingStyleName");
                if (!numStyle.isEmpty())
                    return;
            }
            catch (Exception e) {
                e.printStackTrace(System.err);
            }

            XPropertySet xContentProps = UnoRuntime
                    .queryInterface(XPropertySet.class, xContent);

            xContentProps.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);

            xCursorProps.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER);
            xCursorProps.setPropertyValue("ParaFirstLineIndent", 0);
            xCursorProps.setPropertyValue("ParaIsAutoFirstLineIndent", false);
            xCursorProps.setPropertyValue("ParaLeftMargin", 0);
            xCursorProps.setPropertyValue("ParaRightMargin", 0);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private boolean hasTextAroundFormula(XTextRange xRange, XTextDocument xDoc) {
        XParagraphCursor xParCursor = UnoRuntime
                .queryInterface(XParagraphCursor.class,
                        xRange.getText().createTextCursorByRange(xRange));

        xParCursor.gotoStartOfParagraph(false);
        xParCursor.gotoRange(xRange.getStart(), true);

        if (!xParCursor.getString().trim().isEmpty())
            return true;

        xParCursor.gotoRange(xRange.getEnd(), false);
        xParCursor.gotoEndOfParagraph(true);

        if (!xParCursor.getString().trim().isEmpty())
            return true;

        return false;
    }
}
