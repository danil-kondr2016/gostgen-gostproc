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

import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import org.apache.commons.text.StringSubstitutor;
import org.apache.commons.text.lookup.StringLookup;

/**
 * Обработчик подстановки строк. Осуществляет подстановку на месте макросов
 * вида %...%.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class StringMacroSubstitutor implements MacroSubstitutor.Substitutor {
    private final StringLookup lookup;

    public StringMacroSubstitutor(StringLookup lookup) {
        this.lookup = lookup;
    }
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange) {
        StringSubstitutor substitutor = new StringSubstitutor(lookup, "%", "%", '%');

        String value = substitutor.replace(xRange.getString());
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        xDoc.getText().insertString(xCursor, value, true);
    }
}
