package ru.danilakondr.templater.processing;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexReplace;
import com.sun.star.style.NumberingType;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;

import java.util.concurrent.atomic.AtomicInteger;

/**
 * Обрабатывает стили списков.
 *
 * @author Данила А. Кондратенко
 * @since 0.2.7
 */
public class NumberingStyleProcessor extends Processor {

    public NumberingStyleProcessor(XTextDocument xDoc) {
        super(xDoc);
    }

    @Override
    public void process() throws Exception {
        System.out.println("Processing numbering styles of paragraphs...");
        XEnumerationAccess xEnumAccess = UnoRuntime
                .queryInterface(XEnumerationAccess.class, xDoc.getText());
        XEnumeration xEnum = xEnumAccess.createEnumeration();

        AtomicInteger i = new AtomicInteger(0);
        while (xEnum.hasMoreElements()) {
            System.out.println("Processing numbering style of paragraph #" + i.incrementAndGet() + "...");
            XTextContent xParagraph = UnoRuntime
                    .queryInterface(XTextContent.class, xEnum.nextElement());

            processSingleParagraph(xParagraph);
        }

    }

    private void processSingleParagraph(XTextContent xParagraph) throws  Exception {
        XPropertySet xParProp = UnoRuntime
                .queryInterface(XPropertySet.class, xParagraph);

        try {
            String styleName = (String) xParProp.getPropertyValue("NumberingStyleName");
            if (styleName == null || styleName.isEmpty() || styleName.compareTo("Outline") == 0)
                return;

            XIndexReplace xRules = UnoRuntime
                    .queryInterface(XIndexReplace.class,
                            xParProp.getPropertyValue("NumberingRules"));

            processRules(xRules);
            xParProp.setPropertyValue("NumberingRules", xRules);
        }
        // Значит, это не совсем абзац...
        catch (UnknownPropertyException ignored) {}
    }

    private void processRules(XIndexReplace xRules) throws Exception {
        for (int i = 0; i < 4; i++) {
            PropertyValue[] levelProps = (PropertyValue[])xRules.getByIndex(i);
            processSingleLevel(i, levelProps);
            xRules.replaceByIndex(i, levelProps);
        }
    }

    private void processSingleLevel(int i, PropertyValue[] levelProps) {
        boolean isBullet = false;

        for (int j = 0; j < levelProps.length; j++) {
            if (levelProps[j].Name.compareTo("NumberingType") == 0)
                if (NumberingType.CHAR_SPECIAL == (Short)levelProps[j].Value)
                    isBullet = true;

            if (levelProps[j].Name.compareTo("ParentNumbering") == 0)
                levelProps[j].Value = (short)(i+1);
            if (levelProps[j].Name.compareTo("IndentAt") == 0)
                levelProps[j].Value = 0;
            if (levelProps[j].Name.compareTo("FirstLineIndent") == 0)
                levelProps[j].Value = 1250;
            if (levelProps[j].Name.compareTo("LeftMargin") == 0)
                levelProps[j].Value = 0;
            if (levelProps[j].Name.compareTo("PositionAndSpaceMode") == 0)
                levelProps[j].Value = PositionAndSpaceMode.LABEL_ALIGNMENT;
            if (levelProps[j].Name.compareTo("LabelFollowedBy") == 0)
                levelProps[j].Value = LabelFollow.SPACE;

            if (!isBullet && levelProps[j].Name.compareTo("Prefix") == 0)
                levelProps[j].Value = "";
            if (!isBullet && levelProps[j].Name.compareTo("Suffix") == 0)
                levelProps[j].Value = ")";
            if (!isBullet && levelProps[j].Name.compareTo("ListFormat") == 0)
                levelProps[j].Value = listFormat(i);

            if (isBullet && levelProps[j].Name.compareTo("BulletChar") == 0)
                levelProps[j].Value = "\u2014";
        }
    }

    private String listFormat(int i) {
        return String.format("%%%d%%)", i+1);
    }
}
