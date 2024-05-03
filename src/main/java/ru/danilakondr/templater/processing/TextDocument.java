/*
 * Copyright (c) 2024 Danila A. Kondratenko
 *
 * This file is a part of UNO Templater.
 *
 * UNO Templater is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * UNO Templater is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with UNO Templater.  If not, see <https://www.gnu.org/licenses/>.
 */

package ru.danilakondr.templater.processing;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.*;
import com.sun.star.document.XEmbeddedObjectSupplier2;
import com.sun.star.embed.EmbedUpdateModes;
import com.sun.star.embed.XEmbeddedObject;
import com.sun.star.text.*;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import ru.danilakondr.templater.progress.ProgressInformer;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

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
     * GUID типа объекта формул.
     */
    private static final String MATH_FORMULA_GUID = "078B7ABA-54FC-457F-8551-6147e776a997";
    /**
     * Список всех формул в документе. Должен инициализироваться при первом
     * вызове метода processFormulas().
     */
    private final HashMap<String, Object> formulas;
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
        this.formulas = new HashMap<>();
        this.tables = new HashMap<>();
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
     * Сканирует все таблицы в документе и выбирает те из них, которые не
     * попадают в секции, начинающиеся с <code>eq:</code>
     *
     * @see TextDocument#tables
     * @see TextDocument#scanAllFormulas
     */
    private void scanAllTables() throws Exception
    {
        if (!tables.isEmpty())
            return;

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
    public void processParagraphs(ObjectProcessor<XTextContent> processor, ProgressInformer progress) throws Exception {
        XEnumerationAccess xEnumAccess = UnoRuntime
                .queryInterface(XEnumerationAccess.class, xDoc.getText());
        XEnumeration xEnum = xEnumAccess.createEnumeration();

        int i = 0;
        while (xEnum.hasMoreElements()) {
            XTextContent xParagraph = UnoRuntime
                    .queryInterface(XTextContent.class, xEnum.nextElement());

            progress.inform(++i, -1);
            processor.process(xParagraph, xDoc);
        }
    }

    /**
     * Обрабатывает все изображения по порядку.
     *
     * @param processor обработчик изображения
     * @param progress счётчик прогресса
     */
    public void processImages(ObjectProcessor<Object> processor, ProgressInformer progress) throws Exception {
        XNameAccess graphicObjects = UnoRuntime
                .queryInterface(XTextGraphicObjectsSupplier.class, xDoc)
                .getGraphicObjects();
        String[] names = graphicObjects.getElementNames();

        int i = 0;
        for (String objId : names) {
            progress.inform(++i, names.length);
            processor.process(graphicObjects.getByName(objId), xDoc);
        }
    }

    /**
     * Обрабатывает все формулы по порядку.
     *
     * @param processor обработчик формулы
     * @param progress счётчик прогресса
     */
    public void processFormulas(ObjectProcessor<Object> processor, ProgressInformer progress) throws Exception {
        if (formulas.isEmpty())
            scanAllFormulas();

        AtomicInteger i = new AtomicInteger(0);
        formulas.forEach((k, v) -> {
            progress.inform(i.incrementAndGet(), formulas.size());
            processor.process(v, xDoc);
        });
    }

    public void processTables(ObjectProcessor<XTextTable> processor, ProgressInformer progress) throws Exception {
        if (tables.isEmpty())
            scanAllTables();

        AtomicInteger i = new AtomicInteger(0);
        tables.forEach((k, v) -> {
            progress.inform(i.incrementAndGet(), tables.size());
            processor.process(v, xDoc);
        });
    }

    /**
     * Обновляет все индексы в документе.
     */
    public void updateAllIndexes(ProgressInformer informer) {
        informer.inform(-1, -1);
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
    }

    public void updateAllIndexes() {
        updateAllIndexes((a, b) -> {});
    }
}
