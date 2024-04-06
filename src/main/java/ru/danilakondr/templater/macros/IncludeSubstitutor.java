package ru.danilakondr.templater.macros;

import com.sun.star.beans.PropertyValue;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;

import java.io.File;
import java.util.regex.Pattern;

public class IncludeSubstitutor implements MacroSubstitutor.Substitutor {
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter) {
        Pattern macroPattern = Pattern.compile("%INCLUDE\\((.*)\\)%");
        String macro = xRange.getString();

        if (!macroPattern.matcher(macro).matches())
            return;

        String include = macroPattern.matcher(xRange.getString()).replaceAll("$1");

        File f = new File(include).getAbsoluteFile();
        if (!f.exists()) {
            System.err.printf("File %s not found. Skipping.\n", f);
            return;
        }

        String url = f.toURI().toString();
        try {
            System.out.printf("Including %s\n", url);

            XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
            xCursor.gotoRange(xRange, true);

            XDocumentInsertable xInsertable = UnoRuntime.queryInterface(XDocumentInsertable.class, xCursor);
            xInsertable.insertDocumentFromURL(url, new PropertyValue[0]);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
