package ru.danilakondr.gostproc;

import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.NumberingType;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.*;
import com.sun.star.container.XNameAccess;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик стилей страниц.
 *
 * @author Данила А. Кондратенко
 * @since 2024.03.27
 */
public class PageStyleProcessor extends Processor {
    private final XNameAccess xPageStyles;

    public PageStyleProcessor(XTextDocument xDoc) {
        super(xDoc);

        XStyleFamiliesSupplier xStyleSup = UnoRuntime.queryInterface(
                XStyleFamiliesSupplier.class,
                xDoc
        );
        XNameAccess xStyleFamilies = xStyleSup.getStyleFamilies();
        try {
            xPageStyles = UnoRuntime.queryInterface(
                    XNameAccess.class,
                    xStyleFamilies.getByName("PageStyles")
            );
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Обрабатывает стили &laquo;Базовый&raquo; и
     * &laquo;Первая страница&raquo;. Первая страница по умолчанию
     * не имеет колонтитула.
     *
     * @see PageStyleProcessor#setPageStyle
     * @throws Exception
     */
    @Override
    public void process() throws Exception {
        setPageStyle(xPageStyles, "Standard", true);
        setPageStyle(xPageStyles, "First Page", false);
    }

    /**
     * Установка стилей одной страницы. По умолчанию устанавливает
     * размер бумаги A4 (210 на 297 мм) и поля:
     * <ul>
     *     <li>левое - 30 мм;</li>
     *     <li>правое - 15 мм;</li>
     *     <li>верхнее и нижнее - 20 мм.</li>
     * </ul>
     * <p>
     * Также при условии <code>footer == true</code> вставляется
     * нижний колонтитул с номером страницы посередине.
     *
     * @param xPageStyles словарь, содержащий стили страниц
     * @param styleName имя стиля страницы
     * @param footer указание на вставку нижнего колонтитула
     */
    private void setPageStyle(XNameAccess xPageStyles,
                              String styleName,
                              boolean footer
    )
    {
        try {
            XStyle xStyle = UnoRuntime.queryInterface(XStyle.class, xPageStyles.getByName(styleName));
            XPropertySet xStyleProp = UnoRuntime.queryInterface(XPropertySet.class, xStyle);

            xStyleProp.setPropertyValue("Size", new com.sun.star.awt.Size(21000, 29700));
            xStyleProp.setPropertyValue("LeftMargin", 3000);
            xStyleProp.setPropertyValue("RightMargin", 1500);
            xStyleProp.setPropertyValue("TopMargin", 2000);
            xStyleProp.setPropertyValue("BottomMargin", 2000);
            xStyleProp.setPropertyValue("FooterIsOn", footer);

            if (footer) {
                XText xFooterText = UnoRuntime
                        .queryInterface(XText.class,
                                xStyleProp.getPropertyValue("FooterText"));
                putPageNumber(xFooterText);
            }
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Добавляет номер страницы в область текста.
     *
     * @param xFooterText область текста
     * @throws Exception
     */
    private void putPageNumber(XText xFooterText) throws Exception {
        XTextCursor xCursor = xFooterText.createTextCursorByRange(xFooterText.getStart());

        xCursor.gotoStart(false);
        xCursor.gotoEnd(true);
        xFooterText.insertString(xCursor, "", true);

        xFooterText.insertString(xCursor, "\t", false);

        XTextField xPageNumber = createPageNumber();
        xFooterText.insertTextContent(xCursor, xPageNumber, false);
    }

    /**
     * Создаёт объект номера страницы (текстовое поле).
     *
     * @return объект номера страницы, который можно вставить в любое место
     * @throws Exception
     */
    private XTextField createPageNumber() throws Exception {
        XMultiServiceFactory xMSF = UnoRuntime
                .queryInterface(XMultiServiceFactory.class, xDoc);

        Object oField = xMSF.createInstance("com.sun.star.text.textfield.PageNumber");
        XTextField xField = UnoRuntime.queryInterface(XTextField.class, oField);
        XPropertySet xFieldProp = UnoRuntime.queryInterface(XPropertySet.class, xField);

        xFieldProp.setPropertyValue("NumberingType", NumberingType.ARABIC);

        return xField;
    }
}
