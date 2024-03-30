package ru.danilakondr.gostproc;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.style.LineSpacing;
import com.sun.star.style.LineSpacingMode;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.beans.XPropertySet;

/**
 * Установщик стилей абзацев.
 *
 * @author Данила А. Кондратенко
 * @since 0.1
 */
public class ParagraphStyleProcessor extends Processor {
    private final XNameAccess xParagraphStyles;

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

    @Override
    public void process() throws Exception {
        firstIndent("Text body");
        alignCenter("Caption");
        alignCenter("Drawing");
        alignCenter("Illustration");
        alignCenter("Figure");
        alignLeft("Contents 1");
        alignLeft("Contents 2");
        alignLeft("Contents 3");
    }

    private void defaultParameters(XPropertySet xStyleProp) throws Exception {
        xStyleProp.setPropertyValue("ParaLineSpacing", new LineSpacing(LineSpacingMode.PROP, (short) 150));
        xStyleProp.setPropertyValue("CharHeight", 14.0);
        xStyleProp.setPropertyValue("ParaLeftMargin", 0);
        xStyleProp.setPropertyValue("ParaRightMargin", 0);
        xStyleProp.setPropertyValue("ParaTopMargin", 0);
        xStyleProp.setPropertyValue("ParaBottomMargin", 0);
        xStyleProp.setPropertyValue("ParaFirstLineIndent", 12500);
        xStyleProp.setPropertyValue("ParaOrphans", 0);
        xStyleProp.setPropertyValue("ParaWidows", 0);
    }

    private void firstIndent(String styleName) throws Exception {
        XPropertySet xStyleProp = UnoRuntime.queryInterface(
                XPropertySet.class,
                xParagraphStyles.getByName(styleName)
        );

        defaultParameters(xStyleProp);
        xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.STRETCH);
    }

    private void alignLeft(String styleName) throws Exception {
        XPropertySet xStyleProp = UnoRuntime.queryInterface(
                XPropertySet.class,
                xParagraphStyles.getByName(styleName)
        );

        defaultParameters(xStyleProp);
        xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.LEFT);
    }

    private void alignCenter(String styleName) throws Exception {
        XPropertySet xStyleProp = UnoRuntime.queryInterface(
                XPropertySet.class,
                xParagraphStyles.getByName(styleName)
        );

        defaultParameters(xStyleProp);
        xStyleProp.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER);
    }
}
