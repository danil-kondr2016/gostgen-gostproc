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
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик формул.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0, 0.1.0
 */
public class MathFormulaFixProcessor implements TextDocument.ObjectProcessor<Object> {
    @Override
    public void process(Object object, XTextDocument xDoc) {
        try {
            XPropertySet xFormulaObject = UnoRuntime
                    .queryInterface(XPropertySet.class, object);
            Object oFormula = xFormulaObject.getPropertyValue("Model");

            XPropertySet xPropertySet = UnoRuntime
                    .queryInterface(XPropertySet.class, oFormula);
            String sFormula = (String)xPropertySet.getPropertyValue("Formula");

            xPropertySet.setPropertyValue("Formula", StarMathFixer.fixFormula(sFormula));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
