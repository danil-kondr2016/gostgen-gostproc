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
