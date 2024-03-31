package ru.danilakondr.gostproc.processing;

import com.sun.star.beans.PropertyValue;
import com.sun.star.container.XIndexReplace;
import com.sun.star.style.NumberingType;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик стилей заголовков. Приводит нумерацию в соответствии
 * с ГОСТом.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.3
 */
public class OutlineStyleProcessor extends Processor {

    public OutlineStyleProcessor(XTextDocument xDoc) {
        super(xDoc);
    }

    /**
     * Устанавливает стиль нумерации заголовков.
     */
    @Override
    public void process() throws Exception {
        XChapterNumberingSupplier xSup = UnoRuntime.queryInterface(
                XChapterNumberingSupplier.class,
                xDoc
        );
        XIndexReplace xOutline = xSup.getChapterNumberingRules();
        for (int i = 0; i < 6; i++) {
            PropertyValue[] oNumberingRule = new PropertyValue[8];

            oNumberingRule[0] = createProperty("ParentNumbering", (short)(i+1));
            oNumberingRule[1] = createProperty("IndentAt", 1250);
            oNumberingRule[2] = createProperty("LeftMargin", 0);
            oNumberingRule[3] = createProperty("SymbolTextDistance", 1250);
            oNumberingRule[4] = createProperty("Adjust", HoriOrientation.LEFT);
            oNumberingRule[5] = createProperty("NumberingType", NumberingType.ARABIC);
            oNumberingRule[6] = createProperty("PositionAndSpaceMode", PositionAndSpaceMode.LABEL_ALIGNMENT);
            oNumberingRule[7] = createProperty("LabelFollowedBy", LabelFollow.SPACE);

            xOutline.replaceByIndex(i, oNumberingRule);
        }
    }

    private PropertyValue createProperty(String name, Object value) {
        PropertyValue result = new PropertyValue();
        result.Name = name;
        result.Value = value;

        return result;
    }
}
