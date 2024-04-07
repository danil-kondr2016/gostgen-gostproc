package ru.danilakondr.templater.macros;

import com.sun.star.container.XIndexAccess;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

/**
 * Обработчик макросов в документе.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class MacroSubstitutor {
    /**
     * Документ
     */
    private final XTextDocument xDoc;

    public MacroSubstitutor(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    /**
     * Функциональный интерфейс-обработчик макросов.
     */
    @FunctionalInterface
    public interface Substitutor {
        void substitute(XTextDocument xDoc, XTextRange xRange, Object parameter);

        default boolean test(XTextRange xRange) {
            return xRange.getString().matches("%(.*?)%");
        }
    }

    /**
     * Ищет и обрабатывает макросы в документе.
     *
     * @param proc обработчик макросов
     * @param parameter параметр
     */
    public MacroSubstitutor substitute(Substitutor proc, Object parameter) throws Exception {
        XSearchable xS = UnoRuntime.queryInterface(XSearchable.class, xDoc);
        XSearchDescriptor xSD = xS.createSearchDescriptor();

        xSD.setSearchString("%(.*?)%");
        xSD.setPropertyValue("SearchRegularExpression", true);

        XIndexAccess xAllFound = xS.findAll(xSD);
        for (int i = 0; i < xAllFound.getCount(); i++) {
            Object oFound = xAllFound.getByIndex(i);
            XTextRange xFound = UnoRuntime.queryInterface(XTextRange.class, oFound);
            if (proc.test(xFound)) {
                proc.substitute(xDoc, xFound, parameter);
            }
        }

        return this;
    }
}
