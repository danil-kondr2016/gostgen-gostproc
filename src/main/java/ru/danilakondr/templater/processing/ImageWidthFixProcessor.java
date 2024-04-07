package ru.danilakondr.templater.processing;

import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

import java.util.function.Consumer;

/**
 * Обработчик изображений. Изменяет размеры изображений таким образом, чтобы
 * они помещались целиком на страницу. Если ширина превышает 16,5 см, то высота
 * подстраивается под ширину. Если высота превышает 27,7 см, то ширина
 * подстраивается под высоту.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0, 0.2.6
 */
public class ImageWidthFixProcessor implements TextDocument.ObjectProcessor<Object> {
    @Override
    public void process(Object oImage, XTextDocument xDoc) {
        try {
            XPropertySet xImage = UnoRuntime
                    .queryInterface(XPropertySet.class, oImage);

            Size actualSize = (Size) xImage.getPropertyValue("ActualSize");

            if (actualSize.Width > 16500) {
                long height = 16500L * actualSize.Height / actualSize.Width;

                if (height <= 27700L) {
                    xImage.setPropertyValue("Width", 16500);
                    xImage.setPropertyValue("Height", Long.valueOf(height).intValue());
                } else {
                    long width = 27700L * actualSize.Width / actualSize.Height;

                    xImage.setPropertyValue("Width", Long.valueOf(width).intValue());
                    xImage.setPropertyValue("Height", 27700);
                }
            }
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
