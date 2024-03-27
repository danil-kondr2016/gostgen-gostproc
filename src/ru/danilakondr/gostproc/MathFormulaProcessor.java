package ru.danilakondr.gostproc;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.document.XEmbeddedObjectSupplier;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MathFormulaProcessor extends Processor {
    public static final String MATH_FORMULA_GUID = "078B7ABA-54FC-457F-8551-6147e776a997";
    private final Pattern left, right;

    public MathFormulaProcessor(XTextDocument xDoc) {
        super(xDoc);

        left = Pattern.compile("#\\s*([*/&|=<>]|cdot|times|div)");
        right = Pattern.compile("([\\\\+\\-/&|=<>]|cdot|times|div|plusminus|minusplus)\\s*#");
    }

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

    private void processFormula(XEmbeddedObjectSupplier oFormulaSup) throws Exception {
        Object oFormula = oFormulaSup.getEmbeddedObject();

        XPropertySet xPropertySet = UnoRuntime
                .queryInterface(XPropertySet.class, oFormula);
        String sFormula = (String)xPropertySet.getPropertyValue("Formula");

        Matcher lMatch = left.matcher(sFormula);
        String sFormula1 = lMatch.replaceAll("# {} $1");

        Matcher rMatch = right.matcher(sFormula1);
        String sFormula2 = rMatch.replaceAll("$1 {} #");

        xPropertySet.setPropertyValue("Formula", sFormula2);
    }
}
