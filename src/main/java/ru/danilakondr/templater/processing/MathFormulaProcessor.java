package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.document.XEmbeddedObjectSupplier;
import com.sun.star.document.XEmbeddedObjectSupplier2;
import com.sun.star.embed.EmbedUpdateModes;
import com.sun.star.embed.XEmbeddedObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.commons.text.StringEscapeUtils;

import java.util.Arrays;
import java.util.HashMap;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Обработчик объектов, которые содержат в себе математические формулы
 * LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class MathFormulaProcessor extends Processor {
    /**
     * GUID встроенного объекта LibreOffice Math Formula.
     *
     * @see MathFormulaProcessor#processFormula
     */
    public static final String MATH_FORMULA_GUID = "078B7ABA-54FC-457F-8551-6147e776a997";
    public static HashMap<String, Object> formulas;

    public MathFormulaProcessor(XTextDocument xDoc) {
        super(xDoc);

        formulas = new HashMap<>();
    }

    /**
     * Ищет формулы в документе. Формулы являются встроенными объектами
     * с соответствующим CLSID, указанным в константе
     * <code>MATH_FORMULA_GUID</code>.
     *
     * @see MathFormulaProcessor#processFormula
     * @see MathFormulaProcessor#MATH_FORMULA_GUID
     */
    @Override
    public void process() throws Exception {
        System.out.println("Processing formulas...");

        XTextEmbeddedObjectsSupplier xEmbObj = UnoRuntime.queryInterface(
                XTextEmbeddedObjectsSupplier.class,
                this.xDoc
        );
        XNameAccess xList = xEmbObj.getEmbeddedObjects();
        String[] aElNames = xList.getElementNames();

        for (String aElName : aElNames) {
            Object oFormula = xList.getByName(aElName);
            XPropertySet xFormulaObject = UnoRuntime.queryInterface(
                    XPropertySet.class,
                    oFormula
            );

            xFormulaObject.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);

            String guid = (String) xFormulaObject.getPropertyValue("CLSID");
            if (guid.equalsIgnoreCase(MATH_FORMULA_GUID)) {
                XEmbeddedObject xExt = UnoRuntime
                        .queryInterface(XEmbeddedObjectSupplier2.class, xFormulaObject)
                        .getExtendedControlOverEmbeddedObject();
                formulas.put(aElName, xFormulaObject);
                xExt.setUpdateMode(EmbedUpdateModes.ALWAYS_UPDATE);
            }
        }

        AtomicInteger i = new AtomicInteger(0);
        formulas.forEach((k, v) -> {
            try {
                System.out.printf("Processing formula %d/%d...\n", i.incrementAndGet(), formulas.size());
                XPropertySet xFormula = UnoRuntime
                        .queryInterface(XPropertySet.class, v);
                processFormula(xFormula);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        });
    }

    /**
     * Обрабочик одного объекта формулы.
     *
     * @param xFormulaObject объект, содержащий в себе формулу
     */
    private void processFormula(XPropertySet xFormulaObject) throws Exception {
        Object oFormula = xFormulaObject.getPropertyValue("Model");

        XPropertySet xPropertySet = UnoRuntime
                .queryInterface(XPropertySet.class, oFormula);
        String sFormula = (String)xPropertySet.getPropertyValue("Formula");

        xPropertySet.setPropertyValue("Formula", StarMathFixer.fixFormula(sFormula));
        // Здесь стоит всё-таки подумать над тем, как правильно
        // обрабатывать шрифты...
    }
}
