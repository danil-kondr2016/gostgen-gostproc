package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.document.XEmbeddedObjectSupplier2;
import com.sun.star.embed.EmbedUpdateModes;
import com.sun.star.embed.XEmbeddedObject;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

import java.util.HashMap;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

/**
 * Класс-обёртка над XTextDocument. Содержит методы для работы с объектами,
 * которые принимают на вход обработчик и счётчик процесса.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.0
 */
public class TextDocument {
    /**
     * Документ.
     */
    private final XTextDocument xDoc;
    /**
     * Сравнитель текстовых отрезков.
     */
    private final XTextRangeCompare xCmp;
    /**
     * GUID типа объекта формул.
     */
    private static final String MATH_FORMULA_GUID = "078B7ABA-54FC-457F-8551-6147e776a997";
    /**
     * Список всех формул в документе. Должен инициализироваться при первом
     * вызове метода processFormulas().
     */
    private final HashMap<String, Object> formulas;
    /**
     * Список всех секций в документе. Должен инициализирваться при первом
     * вызове метода streamSections().
     */
    private final HashMap<String, XTextSection> sections;

    /**
     * Интерфейс-обработчик объектов.
     * @param <T> тип объекта
     */
    @FunctionalInterface
    public interface ObjectProcessor<T> {
        void process(T object, XTextDocument xDoc);
    }

    public TextDocument(XTextDocument xDoc) {
        this.xDoc = xDoc;
        this.xCmp = UnoRuntime
                .queryInterface(XTextRangeCompare.class, xDoc.getText());
        this.formulas = new HashMap<>();
        this.sections = new HashMap<>();
    }

    /**
     * Проверяет, находится ли inner внутри outer.
     *
     * @param inner предполагаемый внутренний отрезок
     * @param outer предполагаемый внешний отрезок
     * @return true, если inner внутри outer; false в противном случае
     */
    private boolean isRangeInside(XTextRange inner, XTextRange outer)
    {
        return xCmp.compareRegionStarts(outer, inner) >= 0
                && xCmp.compareRegionEnds(outer, inner) <= 0;
    }

    /**
     * Сканирует все формулы в документе. Метод вызывается один раз.
     * Это необходимо для того, чтобы LibreOffice не входил в глубокую
     * медитацию при обработке формул (видимо, из-за частых запросов к
     * объектной структуре напрямую).
     * <p>
     * Сохраняет все формулы в <code>formulas</code>.
     *
     * @see TextDocument#formulas
     */
    private void scanAllFormulas() throws Exception
    {
        if (!formulas.isEmpty())
            return;

        XTextEmbeddedObjectsSupplier xEmbObj = UnoRuntime.queryInterface(
                XTextEmbeddedObjectsSupplier.class,
                this.xDoc
        );
        XNameAccess embeddedObjects = xEmbObj.getEmbeddedObjects();
        String[] elementNames = embeddedObjects.getElementNames();

        for (String objId : elementNames) {
            Object oFormula = embeddedObjects.getByName(objId);
            XPropertySet xFormulaObject = UnoRuntime.queryInterface(
                    XPropertySet.class,
                    oFormula
            );

            String guid = (String) xFormulaObject.getPropertyValue("CLSID");
            if (guid.equalsIgnoreCase(MATH_FORMULA_GUID)) {
                XEmbeddedObject xExt = UnoRuntime
                        .queryInterface(XEmbeddedObjectSupplier2.class, xFormulaObject)
                        .getExtendedControlOverEmbeddedObject();
                formulas.put(objId, oFormula);
                xExt.setUpdateMode(EmbedUpdateModes.ALWAYS_UPDATE);
            }
        }
    }

    /**
     * Сканирует все секции в документе. Метод вызывается один раз.
     * Сканирование делается от греха подальше, чтобы не произошло глубокой
     * медитации, как в случае с формулами.
     * <p>
     * Сохраняет все секции в <code>sections</code>.
     *
     * @see TextDocument#sections
     * @see TextDocument#scanAllFormulas
     */
    private void scanAllSections() throws Exception {
        if (!sections.isEmpty())
            return;

        XTextSectionsSupplier xSup = UnoRuntime
                .queryInterface(XTextSectionsSupplier.class, xDoc);
        XNameAccess xSections = xSup.getTextSections();

        for (String objId : xSections.getElementNames()) {
            XTextSection xTextSection = UnoRuntime
                    .queryInterface(XTextSection.class, xSections.getByName(objId));
            this.sections.put(objId, xTextSection);
        }
    }

    /**
     * Обрабатывает все абзацы по порядку.
     *
     * @param processor обработчик абзаца
     * @param progress счётчик прогресса
     */
    public TextDocument processParagraphs(ObjectProcessor<XTextContent> processor, ProgressCounter progress) throws Exception {
        XEnumerationAccess xEnumAccess = UnoRuntime
                .queryInterface(XEnumerationAccess.class, xDoc.getText());
        XEnumeration xEnum = xEnumAccess.createEnumeration();

        progress.setShowTotal(false);
        while (xEnum.hasMoreElements()) {
            XTextContent xParagraph = UnoRuntime
                    .queryInterface(XTextContent.class, xEnum.nextElement());

            progress.next();
            processor.process(xParagraph, xDoc);
        }

        return this;
    }

    /**
     * Обрабатывает все изображения по порядку.
     *
     * @param processor обработчик изображения
     * @param progress счётчик прогресса
     */
    public TextDocument processImages(ObjectProcessor<Object> processor, ProgressCounter progress) throws Exception {
        XNameAccess graphicObjects = UnoRuntime
                .queryInterface(XTextGraphicObjectsSupplier.class, xDoc)
                .getGraphicObjects();
        String[] names = graphicObjects.getElementNames();

        progress.setTotal(names.length);
        for (String objId : names) {
            progress.next();
            processor.process(graphicObjects.getByName(objId), xDoc);
        }

        return this;
    }

    /**
     * Обрабатывает все формулы по порядку.
     *
     * @param processor обработчик формулы
     * @param progress счётчик прогресса
     */
    public TextDocument processFormulas(ObjectProcessor<Object> processor, ProgressCounter progress) throws Exception {
        if (formulas.isEmpty())
            scanAllFormulas();

        progress.setTotal(formulas.size());
        formulas.forEach((k, v) -> {
            progress.next();
            processor.process(v, xDoc);
        });

        return this;
    }

    /**
     * Обрабатывает все таблицы в заданном отрезке текста.
     *
     * @param range отрезок
     * @param processor обработчик таблицы
     * @param progress счётчик прогресса
     */
    public TextDocument processTablesInsideRange(XTextRange range, ObjectProcessor<XTextTable> processor, ProgressCounter progress) throws Exception {
        XTextTablesSupplier xSup = UnoRuntime
                .queryInterface(XTextTablesSupplier.class, xDoc);
        XNameAccess textTables = xSup.getTextTables();
        String[] textTablesNames = textTables.getElementNames();
        HashMap<String, XTextTable> textTablesInRange = new HashMap<>();

        for (String objId : textTablesNames) {
            XTextTable xTable = UnoRuntime
                    .queryInterface(XTextTable.class, textTables.getByName(objId));
            XTextRange xTableRange = xTable.getAnchor();

            if (isRangeInside(xTableRange, range)) {
                textTablesInRange.put(objId, xTable);
                progress.incrementTotal();
            }
        }

        AtomicInteger i = new AtomicInteger(0);
        textTablesInRange.forEach((k, v) -> {
            progress.next();
            processor.process(v, xDoc);
        });

        return this;
    }

    /**
     * Обновляет все индексы в документе.
     */
    public TextDocument updateAllIndexes() {
        XDocumentIndexesSupplier xSup = UnoRuntime
                .queryInterface(XDocumentIndexesSupplier.class, xDoc);
        XIndexAccess xIndexes = xSup.getDocumentIndexes();

        try {
            for (int i = 0; i < xIndexes.getCount(); i++) {
                Object oIndex = xIndexes.getByIndex(i);
                XDocumentIndex xIndex = UnoRuntime
                        .queryInterface(XDocumentIndex.class, oIndex);
                xIndex.update();
            }
        }
        catch (java.lang.Exception e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    /**
     * Возвращает поток с секциями. Необходимо для обработки таблиц
     * внутри секций с названиями на <code>tbl:</code>.
     * @return поток с секциями
     */
    public Stream<XTextSection> streamSections() throws Exception {
        if (sections.isEmpty())
            scanAllSections();

        return sections.values().stream();
    }
}
