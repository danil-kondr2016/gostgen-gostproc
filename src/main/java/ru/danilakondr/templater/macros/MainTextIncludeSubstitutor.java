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

import com.sun.star.beans.PropertyValue;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик макроса %MAIN_TEXT%. Вставляет основной текст на нужном месте.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class MainTextIncludeSubstitutor implements MacroSubstitutor.Substitutor {
    private final String mainTextURL;

    public MainTextIncludeSubstitutor(String mainTextURL) {
        this.mainTextURL = mainTextURL;
    }
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange) {
        try {
            XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
            xCursor.gotoRange(xRange, true);

            XDocumentInsertable xInsertable = UnoRuntime.queryInterface(XDocumentInsertable.class, xCursor);
            xInsertable.insertDocumentFromURL(mainTextURL, new PropertyValue[0]);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public boolean test(XTextRange xRange) {
        return xRange.getString().compareTo("%MAIN_TEXT%") == 0;
    }
}
