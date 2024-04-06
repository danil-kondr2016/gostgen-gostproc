package ru.danilakondr.templater.macros;

import com.sun.star.container.XIndexAccess;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

public class MacroSubstitutor {
    private final XTextDocument xDoc;

    public MacroSubstitutor(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    @FunctionalInterface
    public interface Substitutor {
        public void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter);
    }

    public MacroSubstitutor substitute(Substitutor proc, Object parameter) throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%(.*?)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess xAllFound = xS.findAll(xSD);
        for (int i = 0; i < xAllFound.getCount(); i++) {
            Object oFound = xAllFound.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(XTextRange.class, oFound);
            proc.substitute(xDoc, xFound, parameter);
        }

        return this;
    }
}
