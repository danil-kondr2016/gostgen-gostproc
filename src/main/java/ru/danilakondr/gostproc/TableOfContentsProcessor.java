package ru.danilakondr.gostproc;

import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

/**
 * Обработчик, вставляющий оглавление в документ на месте <code>%TOC%</code>.
 * Заменяет только первое вхождение.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class TableOfContentsProcessor extends Processor {
    private final XText xText;
    private final XParagraphCursor xCursor;

    public TableOfContentsProcessor(XTextDocument xDoc) {
        super(xDoc);

        this.xText = xDoc.getText();
        XTextCursor xTextCursor = this.xText.createTextCursorByRange(this.xText.getStart());
        this.xCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
    }

    /**
     * Метод обработки. Ищет первое вхождение %TOC% и вставляет на его
     * место оглавление.
     *
     * @see TableOfContentsProcessor#putTableOfContents
     */
    @Override
    public void process() throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("^([ \\t]*)%TOC%$");
        xSD.setPropertyValue("SearchRegularExpression", true);

        Object oResult = xS.findFirst(xSD);
        if (oResult == null)
            return;

        XTextRange xRange = UnoRuntime.queryInterface(XTextRange.class, oResult);
        xCursor.gotoRange(xRange.getStart(), false);
        xCursor.gotoRange(xRange.getEnd(), true);
        putTableOfContents(xCursor);
    }

    /**
     * Создаёт объект оглавления.
     * @return объект оглавления (сервис <code>com.sun.star.text.ContentIndex</code>)
     */
    private Object createIndex() throws Exception {
        XMultiServiceFactory xMSF = UnoRuntime
                .queryInterface(XMultiServiceFactory.class, xDoc);

        return xMSF.createInstance("com.sun.star.text.ContentIndex");
    }

    /**
     * Непосредственно создаёт и помещает объект оглавления в указанном месте.
     * @see TableOfContentsProcessor#createIndex
     * @param cursor указатель на место вставки
     */
    private void putTableOfContents(XTextCursor cursor) throws Exception {
        Object oIndex = createIndex();

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
}
