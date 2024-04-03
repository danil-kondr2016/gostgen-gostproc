package ru.danilakondr.templater.processing;

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
import java.util.HashSet;
import java.util.regex.Pattern;

/**
 * Обработчик макросов вида <code>%INCLUDE(.*)%</code>. Включает другие файлы
 * в документ, что позволяет применять различные шаблоны.
 * <p>
 * При обработке документа должен вызываться первым, поскольку включаемые
 * документы тоже могут иметь в себе макросы. Внутренние <code>%INCLUDE(...)%</code>
 * обрабатываются на 16 уровней в глубину. Это ограничение необходимо для того, чтобы
 * избежать заполнения всей памяти и зависания.
 * <p>
 * Также обрабатывается макрос %MAIN_TEXT%: на месте его вставляется
 * главный текст документа.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.5
 */
public class DocumentIncluder extends Processor {
    private static final int INCLUDE_DEPTH_LIMIT = 16;
    private final String mainTextURL;
    private final HashSet<String> notExisting;

    public DocumentIncluder(XTextDocument xDoc, String mainTextURL) {
        super(xDoc);
        this.mainTextURL = mainTextURL;
        this.notExisting = new HashSet<>();
    }

    public void process() throws Exception {
        System.out.println("Processing include directives...");
        insertMainText();

        for (int i = 0; i < INCLUDE_DEPTH_LIMIT; i++) {
            if (!processSingleLevel())
                break;
        }

        if (hasNotProcessedIncludes()) {
            throw new Exception("%INCLUDE% nested too deeply; maximal depth is " + INCLUDE_DEPTH_LIMIT);
        }
    }

    /**
     * Ищет все представленные в тексте <code>%INCLUDE(...)%</code>.
     * @return последовательность всех мест, где встречается макрос
     */
    private XIndexAccess findAllIncludes() throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%INCLUDE\\(.*\\)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        return xS.findAll(xSD);
    }

    private boolean hasNotProcessedIncludes() throws Exception {
        XIndexAccess xFound = findAllIncludes();
        HashSet<String> macros = new HashSet<>();

        for (int i = 0; i < xFound.getCount(); i++) {
            macros.add(UnoRuntime.queryInterface(XTextRange.class, xFound.getByIndex(i)).getString());
        }

        return !macros.isEmpty() && macros.stream().noneMatch(notExisting::contains);
    }

    /**
     * Вставляет главный текст документа в шаблон.
     *
     * @see DocumentIncluder#insertDocument 
     * @since 0.2.0
     */
    private void insertMainText() throws Exception {
        System.out.println("Inserting main text...");   
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%MAIN\\_TEXT%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        Object oFound = xS.findFirst(xSD);
        XTextRange xFound = UnoRuntime.queryInterface(
                XTextRange.class,
                oFound
        );
        insertDocument(mainTextURL, xFound);
    }

    /**
     * Обрабатывает один уровень <code>%INCLUDE(...)%</code>.
     *
     * @see DocumentIncluder#insertDocument
     * @see DocumentIncluder#processSingleInclude 
     * @return наличие <code>%INCLUDE(...)%</code> в документе.
     * @since 0.1.5
     */
    private boolean processSingleLevel() throws Exception {
        XIndexAccess found = findAllIncludes();
        if (found.getCount() == 0)
            return false;

        for (int i = 0; i < found.getCount(); i++) {
            Object oFound = found.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(
                    XTextRange.class,
                    oFound
            );
            processSingleInclude(xFound);
        }

        return true;
    }

    /**
     * Обрабатывает один конкретный <code>%INCLUDE(...)%</code>. Вставляет документ
     * на его месте.
     *
     * @see DocumentIncluder#insertDocument 
     * @param xRange место, где найден макрос <code>%INCLUDE(...)%</code>
     * @since 0.1.5
     */
    private void processSingleInclude(XTextRange xRange) throws Exception {
        Pattern macro = Pattern.compile("%INCLUDE\\((.*)\\)%");
        String include = macro.matcher(xRange.getString()).replaceAll("$1");

        if (notExisting.contains(xRange.getString()))
            return;

        File f = new File(include).getAbsoluteFile();
        if (!f.exists()) {
            System.err.printf("File %s not found. Skipping.\n", f);
            notExisting.add(xRange.getString());
            return;
        }

        String url = f.toURI().toString();
        insertDocument(url, xRange);
    }

    /**
     * Вставляет произвольный документ на произвольном месте.
     *
     * @param url URL-адрес документа
     * @param at место, куда нужно вставить документ
     * @since 0.2.0
     */
    private void insertDocument(String url, XTextRange at) throws Exception {
        System.out.printf("Including %s\n", url);
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(at);
        xCursor.gotoRange(at, true);
        XDocumentInsertable xInsertable = UnoRuntime.queryInterface(XDocumentInsertable.class, xCursor);
        xInsertable.insertDocumentFromURL(url, new PropertyValue[0]);
    }
}
