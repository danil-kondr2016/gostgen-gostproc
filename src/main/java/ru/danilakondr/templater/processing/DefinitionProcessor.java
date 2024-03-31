package ru.danilakondr.templater.processing;

import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;
import com.sun.star.container.XIndexAccess;

import java.io.IOException;
import java.util.regex.Pattern;

public class DefinitionProcessor extends Processor {
    private final Definitions definitions;

    public DefinitionProcessor(XTextDocument xDoc, String varFile) throws IOException {
        super(xDoc);
        this.definitions = new Definitions(varFile);
    }

    @Override
    public void process() throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%(.*?)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess xAllFound = xS.findAll(xSD);
        for (int i = 0; i < xAllFound.getCount(); i++) {
            Object oFound = xAllFound.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(XTextRange.class, oFound);
            processSingleProperty(xFound);
        }
    }

    private void processSingleProperty(XTextRange xRange) throws Exception {
        Pattern macro = Pattern.compile("%(.*?)%");
        String property = macro.matcher(xRange.getString()).replaceAll("$1");

        System.out.println("Found property " + property);

        if (!definitions.containsKey(property))
            return;

        String value = definitions.get(property);

        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        xDoc.getText().insertString(xCursor, value, true);
    }
}
