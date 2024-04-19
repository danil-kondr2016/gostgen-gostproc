package ru.danilakondr.templater.macros;

import com.sun.star.beans.PropertyValue;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;

/**
 * Обработчик макроса %MAIN_TEXT%. Вставляет основной текст на нужном месте.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class MainTextIncludeSubstitutor implements MacroSubstitutor.Substitutor {
    private final String mainTextURL;

    public MainTextIncludeSubstitutor(String mainTextURL) {
        this.mainTextURL = mainTextURL;
    }
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange) {
        try {
            System.out.printf("Including %s\n", mainTextURL);

            XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
            xCursor.gotoRange(xRange, true);

            XDocumentInsertable xInsertable = UnoRuntime.queryInterface(XDocumentInsertable.class, xCursor);
            xInsertable.insertDocumentFromURL(mainTextURL, new PropertyValue[0]);
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
