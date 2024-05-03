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

import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик изображений. Изменяет размеры изображений таким образом, чтобы
 * они помещались целиком на страницу. Если ширина превышает 16,5 см, то высота
 * подстраивается под ширину. Если высота превышает 24 см, то ширина
 * подстраивается под высоту.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0, 0.2.6
 */
public class ImageSizeFixProcessor implements TextDocument.ObjectProcessor<Object> {
    private static final int TEXT_WIDTH = 16500;
    private static final int TEXT_HEIGHT = 23000;
    @Override
    public void process(Object oImage, XTextDocument xDoc) {
        try {
            XPropertySet xImage = UnoRuntime
                    .queryInterface(XPropertySet.class, oImage);


            Size actualSize = (Size) xImage.getPropertyValue("ActualSize");

            xImage.setPropertyValue("Width", actualSize.Width);
            xImage.setPropertyValue("Height", actualSize.Height);

            int w = (Integer)xImage.getPropertyValue("Width");

            if (w > TEXT_WIDTH) {
                long height = ((long) TEXT_WIDTH) * actualSize.Height / actualSize.Width;

                xImage.setPropertyValue("Width", TEXT_WIDTH);
                xImage.setPropertyValue("Height", Long.valueOf(height).intValue());
            }

            int h = (Integer)xImage.getPropertyValue("Height");
            if (h > TEXT_HEIGHT) {
                long width = ((long) TEXT_HEIGHT) * actualSize.Width / actualSize.Height;

                xImage.setPropertyValue("Width", Long.valueOf(width).intValue());
                xImage.setPropertyValue("Height", TEXT_HEIGHT);
            }
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
