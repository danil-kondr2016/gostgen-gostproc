package ru.danilakondr.gostproc;

import com.sun.star.awt.FontSlant;
import com.sun.star.awt.FontWeight;
import com.sun.star.container.XNameAccess;
import com.sun.star.style.*;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.beans.XPropertySet;

/**
 * Установщик стилей абзацев. Оформляет стили абзацев в соответствии с
 * ГОСТом.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class ParagraphStyleProcessor extends Processor {
    private final XNameAccess xParagraphStyles;

    /**
     * Тип выравнивания.
     */
    public enum Indent {
        LEFT, PARAGRAPH, HEADING, CENTER, RIGHT
    }

    public ParagraphStyleProcessor(XTextDocument xDoc) {
        super(xDoc);

        XStyleFamiliesSupplier xStyleSup = UnoRuntime.queryInterface(
                XStyleFamiliesSupplier.class,
                xDoc
        );
        XNameAccess xStyleFamilies = xStyleSup.getStyleFamilies();
        try {
            xParagraphStyles = UnoRuntime.queryInterface(
                    XNameAccess.class,
                    xStyleFamilies.getByName("ParagraphStyles")
            );
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Редактирует стили абзацев.
     */
    @Override
    public void process() throws Exception {
        setStandardParagraphStyle();
        firstIndent("Text body");
        alignCenter("Caption");
        alignCenter("Drawing");
        alignCenter("Illustration");
        alignCenter("Figure");
        alignCenter("Footer");
        alignCenter("Footer left");
        alignCenter("Footer right");
        alignLeft("Contents 1");
        alignLeft("Contents 2");
        alignLeft("Contents 3");
        heading("Heading", Indent.HEADING, false);
        heading("Heading 1", Indent.HEADING, true);
        heading("Heading 2", Indent.HEADING, true);
        heading("Heading 3", Indent.HEADING, true);
        heading("Heading 4", Indent.HEADING, true);
        heading("Contents Heading", Indent.CENTER, false);
    }

    private void setStandardParagraphStyle() throws Exception {
        XPropertySet xStyleProp = getStyleProperties("Standard");
        xStyleProp.setPropertyValue("CharHeight", 14.0f);
    }

    private XPropertySet getStyleProperties(String styleName) throws Exception {
        return UnoRuntime.queryInterface(
                XPropertySet.class,
                xParagraphStyles.getByName(styleName)
        );
    }

    /**
     * Редактирует стиль заголовка.
     *
     * @param styleName название стиля заголовка
     * @param indent тип выравнивания
     * @param numbered установка нумерации
     * @since 0.1.2
     */
    private void heading(String styleName, Indent indent, boolean numbered) throws Exception {
        XPropertySet xStyleProp = getStyleProperties(styleName);

        setParagraphStyleParameters(xStyleProp, indent, true);

        /* Нумерация */
        if (numbered)
            xStyleProp.setPropertyValue("NumberingStyleName", "Outline");
        else
            xStyleProp.setPropertyValue("NumberingStyleName", null);
    }

    private void setParagraphStyleParameters(XPropertySet xStyleProp,
                                             Indent indent,
                                             boolean bold)
            throws Exception
    {
        xStyleProp.setPropertyValue("ParaLineSpacing", new LineSpacing(LineSpacingMode.PROP, (short) 150));
        xStyleProp.setPropertyValue("ParaLeftMargin", 0);
        xStyleProp.setPropertyValue("ParaRightMargin", 0);
        xStyleProp.setPropertyValue("ParaTopMargin", 0);
        xStyleProp.setPropertyValue("ParaBottomMargin", 0);
        xStyleProp.setPropertyValue("ParaOrphans", (byte)0);
        xStyleProp.setPropertyValue("ParaWidows", (byte)0);

        xStyleProp.setPropertyValue("CharHeight", 14.0f);
        xStyleProp.setPropertyValue("CharFontName", "Liberation Serif");

        xStyleProp.setPropertyValue("CharWeight", bold ? FontWeight.BOLD : FontWeight.NORMAL);
        xStyleProp.setPropertyValue("CharPosture", FontSlant.NONE);

        xStyleProp.setPropertyValue("ParaFirstLineIndent", 0);
        switch (indent) {
            case HEADING:
                xStyleProp.setPropertyValue("ParaFirstLineIndent", 1250);
            case LEFT:
                xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.LEFT);
                break;
            case PARAGRAPH:
                xStyleProp.setPropertyValue("ParaFirstLineIndent", 1250);
                xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.STRETCH);
                break;
            case RIGHT:
                xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.RIGHT);
                break;
            case CENTER:
                xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER);
                break;
        }
    }

    private void firstIndent(String styleName) throws Exception {
        XPropertySet xStyleProp = getStyleProperties(styleName);

        setParagraphStyleParameters(xStyleProp, Indent.PARAGRAPH, false);
    }

    private void alignLeft(String styleName) throws Exception {
        XPropertySet xStyleProp = UnoRuntime.queryInterface(
                XPropertySet.class,
                xParagraphStyles.getByName(styleName)
        );

        setParagraphStyleParameters(xStyleProp, Indent.LEFT, false);
    }

    private void alignCenter(String styleName) throws Exception {
        XPropertySet xStyleProp = UnoRuntime.queryInterface(
                XPropertySet.class,
                xParagraphStyles.getByName(styleName)
        );

        setParagraphStyleParameters(xStyleProp, Indent.CENTER, false);
    }
}
