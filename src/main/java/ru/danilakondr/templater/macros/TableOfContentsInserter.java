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

package ru.danilakondr.templater.macros;

import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик, вставляющий оглавление в документ на месте <code>%TOC%</code>.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class TableOfContentsInserter implements MacroSubstitutor.Substitutor {
    /**
     * Создаёт объект оглавления.
     * @return объект оглавления (сервис <code>com.sun.star.text.ContentIndex</code>)
     */
    private Object createIndex(XTextDocument xDoc) throws Exception {
        XMultiServiceFactory xMSF = UnoRuntime
                .queryInterface(XMultiServiceFactory.class, xDoc);

        return xMSF.createInstance("com.sun.star.text.ContentIndex");
    }

    /**
     * Непосредственно создаёт и помещает объект оглавления в указанном месте.
     * @see TableOfContentsInserter#createIndex
     * @param cursor указатель на место вставки
     */
    private void putTableOfContents(XTextDocument xDoc, XText xText, XTextCursor cursor) throws Exception {
        Object oIndex = createIndex(xDoc);

        XDocumentIndex xIndex = UnoRuntime
                .queryInterface(XDocumentIndex.class, oIndex);
        XPropertySet xIndexProp = UnoRuntime
                .queryInterface(XPropertySet.class, xIndex);

        xText.insertTextContent(cursor, xIndex, true);

        // Сначала добавить и только потом выставлять свойства!
        xIndexProp.setPropertyValue("CreateFromOutline", true);
        xIndexProp.setPropertyValue("Title", "Оглавление");

        xIndex.update();
    }

    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange) {
        XText xText = xRange.getText();
        XTextCursor xCursor = xText.createTextCursorByRange(xRange);

        xCursor.gotoRange(xRange, true);
        try {
            putTableOfContents(xDoc, xText, xCursor);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public boolean test(XTextRange xRange) {
        return xRange.getString().contains("%TOC%");
    }
}
