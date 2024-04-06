package ru.danilakondr.templater.macros;

import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import org.apache.commons.text.StringSubstitutor;

public class StringMacroSubstitutor implements MacroSubstitutor.Substitutor {
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter) {
        String macro = xRange.getString().trim();
        StringSubstitutor substitutor = (StringSubstitutor) parameter;
        boolean containsKey = substitutor.getStringLookup().lookup(macro.replaceAll("%(.*?)%", "$1")) != null;
        if (!containsKey)
            return;

        System.out.printf("Processing macro %s...\n", xRange.getString());

        String value = substitutor.replace(xRange.getString());
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        xDoc.getText().insertString(xCursor, value, true);
    }
}
