package ru.danilakondr.templater.processing;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.*;
import com.sun.star.document.XEmbeddedObjectSupplier2;
import com.sun.star.embed.EmbedUpdateModes;
import com.sun.star.embed.XEmbeddedObject;
import com.sun.star.text.*;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
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

    private final HashMap<String, XTextTable> tables;

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
        this.tables = new HashMap<>();
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
        int a = xCmp.compareRegionStarts(outer, inner);
        int b = xCmp.compareRegionEnds(outer, inner);
        System.err.printf("  @C %+d %+d %n", a, b);
        return a >= 0 && b <= 0;
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
     * Сканирует все таблицы в документе и выбирает те из них, которые не
     * попадают в секции, начинающиеся с <code>eq:</code>
     *
     * @see TextDocument#tables
     * @see TextDocument#scanAllFormulas
     * @see TextDocument#scanAllSections
     */
    private void scanAllTables() throws Exception
    {
        if (!tables.isEmpty())
            return;
        if (sections.isEmpty())
            scanAllSections();

        XTextTablesSupplier xSup = UnoRuntime
                .queryInterface(XTextTablesSupplier.class, xDoc);
        XNameAccess textTables = xSup.getTextTables();
        String[] textTablesNames = textTables.getElementNames();
        HashMap<String, XTextTable> allTables = new HashMap<>();

        for (String objId : textTablesNames) {
            XTextTable xTable = UnoRuntime
                    .queryInterface(XTextTable.class, textTables.getByName(objId));
            allTables.put(objId, xTable);
        }

        HashSet<String> toRemove = new HashSet<>();
        for (Map.Entry<String, XTextTable> f : allTables.entrySet()) {
            XTextRange xTableRange = f.getValue().getAnchor();
            XPropertySet xTableRangeProp = UnoRuntime
                    .queryInterface(XPropertySet.class, xTableRange);
            XTextSection xSection = UnoRuntime
                    .queryInterface(XTextSection.class,
                            xTableRangeProp.getPropertyValue("TextSection"));
            if (xSection == null)
                continue;

            XNamed xSectionName = UnoRuntime
                    .queryInterface(XNamed.class, xSection);
            if (xSectionName.getName().startsWith("eq:"))
                toRemove.add(f.getKey());
        }

        for (String x : allTables.keySet()) {
            if (toRemove.contains(x))
                continue;
            tables.put(x, allTables.get(x));
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

    public TextDocument processTables(ObjectProcessor<XTextTable> processor, ProgressCounter progress) throws Exception {
        if (tables.isEmpty())
            scanAllTables();

        progress.setTotal(tables.size());
        tables.forEach((k, v) -> {
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
}
