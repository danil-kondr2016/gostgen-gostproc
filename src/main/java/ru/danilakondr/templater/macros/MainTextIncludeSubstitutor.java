package ru.danilakondr.templater.macros;

import com.sun.star.beans.PropertyValue;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;

import java.util.regex.Pattern;

public class MainTextIncludeSubstitutor implements MacroSubstitutor.Substitutor {
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter) {
        Pattern macroPattern = Pattern.compile("%MAIN_TEXT%");
        String macro = xRange.getString();

        if (!macroPattern.matcher(macro).matches())
            return;

        String include = (String)parameter;
        try {
            System.out.printf("Including %s\n", include);

            XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
            xCursor.gotoRange(xRange, true);

            XDocumentInsertable xInsertable = UnoRuntime.queryInterface(XDocumentInsertable.class, xCursor);
            xInsertable.insertDocumentFromURL(include, new PropertyValue[0]);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
