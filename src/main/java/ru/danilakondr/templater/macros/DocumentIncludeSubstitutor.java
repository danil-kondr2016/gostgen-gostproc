package ru.danilakondr.templater.macros;

import com.sun.star.beans.PropertyValue;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.regex.Pattern;

/**
 * Обрабатывает макросы вида %INCLUDE(...)%. Вставляет содержимое заданных
 * файлов на нужные места. Обработка %INCLUDE(...)% может происходить несколько
 * раз.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0, 0.1.5
 */
public class DocumentIncludeSubstitutor implements MacroSubstitutor.Substitutor {
    private static final Pattern macroPattern = Pattern.compile("%INCLUDE\\((.*)\\)%");
    @Override
    public void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter) {
        String include = macroPattern.matcher(xRange.getString()).replaceAll("$1");

        File f = new File(include).getAbsoluteFile();
        if (!f.exists()) {
            throw new RuntimeException(new FileNotFoundException(f.getAbsolutePath()));
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

    @Override
    public boolean test(XTextRange xRange) {
        return macroPattern.matcher(xRange.getString()).matches();
    }
}
