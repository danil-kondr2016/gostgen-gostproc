package ru.danilakondr.templater.processing;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexReplace;
import com.sun.star.style.NumberingType;
import com.sun.star.text.LabelFollow;
import com.sun.star.text.PositionAndSpaceMode;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

import java.util.function.Consumer;

/**
 * Обработчик стилей нумерации. Изначально был отдельным процессором, который
 * содержал в себе документ, однако позже реализован в виде функционального
 * класса-обработчика.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0, 0.2.7
 */
public class NumberingStyleProcessor implements TextDocument.ObjectProcessor<XTextContent> {
    @Override
    public void process(XTextContent xParagraph, XTextDocument xDoc) {
        XPropertySet xParProp = UnoRuntime
                .queryInterface(XPropertySet.class, xParagraph);

        try {
            String styleName = (String) xParProp.getPropertyValue("NumberingStyleName");
            if (styleName == null || styleName.isEmpty() || styleName.compareTo("Outline") == 0)
                return;

            XIndexReplace xRules = UnoRuntime
                    .queryInterface(XIndexReplace.class,
                            xParProp.getPropertyValue("NumberingRules"));

            for (int i = 0; i < 4; i++) {
                PropertyValue[] levelProps = (PropertyValue[])xRules.getByIndex(i);
                processSingleLevel(i, levelProps);
                xRules.replaceByIndex(i, levelProps);
            }
            xParProp.setPropertyValue("NumberingRules", xRules);
        }
        // Значит, это не совсем абзац...
        catch (UnknownPropertyException ignored) {}
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void processSingleLevel(int levelIndex, PropertyValue[] levelProps) {
        boolean isBullet = false;

        for (int j = 0; j < levelProps.length; j++) {
            if (levelProps[j].Name.compareTo("NumberingType") == 0)
                if (NumberingType.CHAR_SPECIAL == (Short)levelProps[j].Value)
                    isBullet = true;

            if (levelProps[j].Name.compareTo("ParentNumbering") == 0)
                levelProps[j].Value = (short)(levelIndex+1);
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
                levelProps[j].Value = listFormat(levelIndex);

            if (isBullet && levelProps[j].Name.compareTo("BulletChar") == 0)
                levelProps[j].Value = "\u2014";
        }
    }

    private String listFormat(int levelIndex) {
        return String.format("%%%d%%)", levelIndex+1);
    }
}
