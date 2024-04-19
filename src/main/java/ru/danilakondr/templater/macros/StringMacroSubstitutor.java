package ru.danilakondr.templater.macros;

import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import org.apache.commons.text.StringSubstitutor;
import org.apache.commons.text.lookup.StringLookup;

/**
 * Обработчик подстановки строк. Осуществляет подстановку на месте макросов
 * вида %...%.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class StringMacroSubstitutor implements MacroSubstitutor.Substitutor {
    private final StringLookup lookup;

    public StringMacroSubstitutor(StringLookup lookup) {
        this.lookup = lookup;
    }
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange) {
        StringSubstitutor substitutor = new StringSubstitutor(lookup, "%", "%", '%');

        String value = substitutor.replace(xRange.getString());
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        xDoc.getText().insertString(xCursor, value, true);
    }
}
