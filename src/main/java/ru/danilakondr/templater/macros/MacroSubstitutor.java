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

import com.sun.star.container.XIndexAccess;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

/**
 * Обработчик макросов в документе.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class MacroSubstitutor {
    /**
     * Документ
     */
    private final XTextDocument xDoc;

    public MacroSubstitutor(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    /**
     * Функциональный интерфейс-обработчик макросов.
     */
    @FunctionalInterface
    public interface Substitutor {
        void substitute(XTextDocument xDoc, XTextRange xRange);

        default boolean test(XTextRange xRange) {
            return xRange.getString().matches("%(.*?)%");
        }
    }

    /**
     * Ищет и обрабатывает макросы в документе.
     *
     * @param proc обработчик макросов
     */
    public void substitute(Substitutor proc) throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%(.*?)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess xAllFound = xS.findAll(xSD);
        for (int i = 0; i < xAllFound.getCount(); i++) {
            Object oFound = xAllFound.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(XTextRange.class, oFound);
            if (proc.test(xFound)) {
                proc.substitute(xDoc, xFound);
            }
        }
    }
}
