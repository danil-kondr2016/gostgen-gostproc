/*
 * Copyright (c) 2024 Danila A. Kondratenko
 *
 * This file is a part of UNO Templater.
 *
 * UNO Templater is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * UNO Templater is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with UNO Templater.  If not, see <https://www.gnu.org/licenses/>.
 */

package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;

/**
 * Выравниватель объектов.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.2
 */
public class SingleObjectAligner implements TextDocument.ObjectProcessor<Object> {
    @Override
    public void process(Object object, XTextDocument xDoc) {
        try {
            XTextContent xContent = UnoRuntime
                    .queryInterface(XTextContent.class, object);
            XTextRange xRange = xContent.getAnchor();

            if (xRange.getText() != xDoc.getText())
                return;

            XTextCursor xCursor = xRange.getText().createTextCursorByRange(xRange);
            if (hasTextAroundObject(xRange))
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

    private boolean hasTextAroundObject(XTextRange xRange) {
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
