package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.document.XEmbeddedObjectSupplier;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;
import org.apache.commons.text.StringEscapeUtils;

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

    public MathFormulaProcessor(XTextDocument xDoc) {
        super(xDoc);
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
        XTextEmbeddedObjectsSupplier xEmbObj = UnoRuntime.queryInterface(
                XTextEmbeddedObjectsSupplier.class,
                this.xDoc
        );
        XNameAccess xList = xEmbObj.getEmbeddedObjects();
        String[] aElNames = xList.getElementNames();

        for (String aElName : aElNames) {
            Object oFormula = xList.getByName(aElName);
            XPropertySet xFormula = UnoRuntime.queryInterface(
                    XPropertySet.class,
                    oFormula
            );
            String guid = (String) xFormula.getPropertyValue("CLSID");
            if (guid.equalsIgnoreCase(MATH_FORMULA_GUID)) {
                XEmbeddedObjectSupplier xObjSup = UnoRuntime
                        .queryInterface(XEmbeddedObjectSupplier.class, xFormula);
                processFormula(xObjSup);
            }
        }
    }

    /**
     * Обрабочик одного объекта формулы.
     *
     * @param oFormulaSup объект, содержащий в себе формулу
     */
    private void processFormula(XEmbeddedObjectSupplier oFormulaSup) throws Exception {
        Object oFormula = oFormulaSup.getEmbeddedObject();

        XPropertySet xPropertySet = UnoRuntime
                .queryInterface(XPropertySet.class, oFormula);
        String sFormula = (String)xPropertySet.getPropertyValue("Formula");

        System.out.printf("Processing formula \"%s\"...\n",
                StringEscapeUtils.escapeJava(sFormula));

        xPropertySet.setPropertyValue("Formula", StarMathFixer.fixFormula(sFormula));
        // Здесь стоит всё-таки подумать над тем, как правильно
        // обрабатывать шрифты...
    }
}
