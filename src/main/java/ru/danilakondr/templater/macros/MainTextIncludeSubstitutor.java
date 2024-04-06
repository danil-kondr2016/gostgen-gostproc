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

    @Override
    public boolean test(XTextRange xRange) {
        return xRange.getString().compareTo("%MAIN_TEXT%") == 0;
    }
}
