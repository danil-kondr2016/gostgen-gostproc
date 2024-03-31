package ru.danilakondr.templater.processing;

import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;
import com.sun.star.container.XIndexAccess;

import java.io.IOException;
import java.util.Set;

import org.apache.commons.text.StringSubstitutor;

/**
 * Класс, обрабатывающий макросы, разворачивающиеся в строки.
 *
 * @author Данила А. Кондратенко
 * @since 0.2.1
 */
public class MacroProcessor extends Processor {
    private final Macros macros;
    private final StringSubstitutor substitutor;
    private static final Set<String> FORBIDDEN_MACROS = Set.of("%TOC%", "%MAIN_TEXT%");

    /**
     * Определяет, является ли макрос запрещённым для строковой подстановки.
     * Запрещённые макросы обрабатываются отдельно другими классами:
     *
     * <ul>
     *     <li><code>%INCLUDE(...)%</code>, <code>%MAIN_TEXT%</code>: <code>DocumentIncluder</code></li>
     *     <li><code>%TOC%</code>: <code>TableOfContentsInserter<code/></li>
     * </ul>
     * @param macro макрос, нуждающийся в проверке
     * @return является ли макрос запрещённым
     * @since 0.2.2
     */
    private static boolean isForbiddenMacro(String macro) {
        if (FORBIDDEN_MACROS.contains(macro))
            return true;

        if (macro.matches("%INCLUDE\\((.*)\\)%"))
            return true;

        return false;
    }

    public MacroProcessor(XTextDocument xDoc, String varFile) throws IOException {
        super(xDoc);
        this.macros = new Macros(varFile);
        this.substitutor = new StringSubstitutor(this.macros);
        this.substitutor.setVariablePrefix('%').setVariableSuffix('%');
    }

    @Override
    public void process() throws Exception {
        this.macros.forEach((k, v) -> {
            System.err.printf("%s: %s\n", k, v);
        });

        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%(.*?)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess xAllFound = xS.findAll(xSD);
        for (int i = 0; i < xAllFound.getCount(); i++) {
            Object oFound = xAllFound.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(XTextRange.class, oFound);
            processSingleProperty(xFound);
        }
    }

    /**
     * Обрабатывает один макрос в заданном месте.
     *
     * @param xRange место, где находится макрос
     * @since 0.2.1
     */
    private void processSingleProperty(XTextRange xRange) throws Exception {
        String macro = xRange.getString().trim();
        boolean containsKey = macros.containsKey(macro.replaceAll("%(.*?)%", "$1"));

        if (!containsKey && !isForbiddenMacro(macro)) {
            System.err.printf("Macro %s has not been specified, skipping\n", macro);
            return;
        }
        else if (containsKey && isForbiddenMacro(macro)) {
            System.err.printf("Macro %s cannot be overridden, skipping\n", macro);
            return;
        }


        String value = substitutor.replace(xRange.getString());
        XTextCursor xCursor = xDoc.getText().createTextCursorByRange(xRange);
        xCursor.gotoRange(xRange, true);
        xDoc.getText().insertString(xCursor, value, true);
    }
}
