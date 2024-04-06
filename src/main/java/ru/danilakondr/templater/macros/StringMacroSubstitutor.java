package ru.danilakondr.templater.macros;

import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import org.apache.commons.text.StringSubstitutor;
import org.apache.commons.text.lookup.StringLookup;

public class StringMacroSubstitutor implements MacroSubstitutor.Substitutor {
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter) {
        StringLookup lookup = (StringMacros)parameter;
        StringSubstitutor substitutor = new StringSubstitutor(lookup, "%", "%", '%');

        String value = substitutor.replace(xRange.getString());
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        xDoc.getText().insertString(xCursor, value, true);
    }
}
