package ru.danilakondr.gostproc;

import com.sun.star.beans.PropertyValue;
import com.sun.star.container.XIndexAccess;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

import java.io.File;
import java.util.regex.Pattern;

/**
 * Обработчик макросов вида <code>%INCLUDE(.*)%</code>. Включает другие файлы
 * в документ, что позволяет применять различные шаблоны.
 * <p>
 * При обработке документа должен вызываться первым, поскольку включаемые
 * документы тоже могут иметь в себе макросы. Внутренние <code>%INCLUDE(...)%</code>
 * не обрабатываются.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.5
 */
public class DocumentIncluder extends Processor {
    public DocumentIncluder(XTextDocument xDoc) {
        super(xDoc);
    }

    @Override
    public void process() throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%INCLUDE\\(.*\\)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess found = xS.findAll(xSD);
        for (int i = 0; i < found.getCount(); i++) {
            Object oFound = found.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(
                    XTextRange.class,
                    oFound
            );
            processSingleInclude(xFound);
        }
    }

    private void processSingleInclude(XTextRange xRange) throws Exception {
        Pattern macro = Pattern.compile("%INCLUDE\\((.*)\\)%");
        String include = macro.matcher(xRange.getString()).replaceAll("$1");

        System.out.printf("processSingleInclude: %s -> %s\n", xRange.getString(), include);

        File f = new File(include).getAbsoluteFile();
        if (!f.exists()) {
            System.err.printf("File %s not found. Skipping.\n", f);
            return;
        }

        String url = f.toURI().toString();
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        XDocumentInsertable xInsertable = UnoRuntime.queryInterface(XDocumentInsertable.class, xCursor);
        PropertyValue[] p = new PropertyValue[0];
        xInsertable.insertDocumentFromURL(url, p);
    }
}
