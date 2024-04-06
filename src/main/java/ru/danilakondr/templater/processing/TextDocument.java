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
import java.util.function.Consumer;
import java.util.stream.Stream;

public class TextDocument {
    private final XTextDocument xDoc;
    private static final String MATH_FORMULA_GUID = "078B7ABA-54FC-457F-8551-6147e776a997";

    @FunctionalInterface
    public interface ObjectProcessor<T> {
        void process(T object, XTextDocument xDoc);
    }

    public TextDocument(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    public TextDocument processParagraphs(ObjectProcessor<XTextContent> proc, ProgressCounter progress) throws Exception {
        XEnumerationAccess xEnumAccess = UnoRuntime
                .queryInterface(XEnumerationAccess.class, xDoc.getText());
        XEnumeration xEnum = xEnumAccess.createEnumeration();

        progress.setShowTotal(false);
        while (xEnum.hasMoreElements()) {
            XTextContent xParagraph = UnoRuntime
                    .queryInterface(XTextContent.class, xEnum.nextElement());

            progress.next();
            proc.process(xParagraph, xDoc);
        }

        return this;
    }

    public TextDocument processImages(ObjectProcessor<Object> proc, ProgressCounter progress) throws Exception {
        XNameAccess graphicObjects = UnoRuntime
                .queryInterface(XTextGraphicObjectsSupplier.class, xDoc)
                .getGraphicObjects();
        String[] names = graphicObjects.getElementNames();

        progress.setTotal(names.length);
        for (String objId : names) {
            progress.next();
            proc.process(graphicObjects.getByName(objId), xDoc);
        }

        return this;
    }

    public TextDocument processFormulas(ObjectProcessor<XPropertySet> proc, ProgressCounter progress) throws Exception {
        XTextEmbeddedObjectsSupplier xEmbObj = UnoRuntime.queryInterface(
                XTextEmbeddedObjectsSupplier.class,
                this.xDoc
        );
        XNameAccess embeddedObjects = xEmbObj.getEmbeddedObjects();
        String[] elementNames = embeddedObjects.getElementNames();
        HashMap<String, Object> formulas = new HashMap<>();

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
                progress.incrementTotal();
            }
        }

        formulas.forEach((k, v) -> {
            progress.next();
            XPropertySet xFormula = UnoRuntime
                    .queryInterface(XPropertySet.class, v);
            proc.process(xFormula, xDoc);
        });

        return this;
    }

    public TextDocument processTablesInsideRange(XTextRange range, ObjectProcessor<XTextTable> proc, ProgressCounter progress) throws Exception {
        XTextTablesSupplier xSup = UnoRuntime
                .queryInterface(XTextTablesSupplier.class, xDoc);
        XNameAccess textTables = xSup.getTextTables();
        String[] textTablesNames = textTables.getElementNames();
        HashMap<String, XTextTable> textTablesInRange = new HashMap<>();

        XTextRangeCompare xCmp = UnoRuntime
                .queryInterface(XTextRangeCompare.class, xDoc.getText());

        for (String objId : textTablesNames) {
            XTextTable xTable = UnoRuntime
                    .queryInterface(XTextTable.class, textTables.getByName(objId));
            XTextRange xTableRange = xTable.getAnchor();

            if (xCmp.compareRegionStarts(range, xTableRange) >= 0 && xCmp.compareRegionEnds(range, xTableRange) <= 0) {
                textTablesInRange.put(objId, xTable);
                progress.incrementTotal();
            }
        }

        AtomicInteger i = new AtomicInteger(0);
        textTablesInRange.forEach((k, v) -> {
            progress.next();
            proc.process(v, xDoc);
        });

        return this;
    }

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

    public Stream<XTextSection> scanSections() throws Exception {
        XTextSectionsSupplier xSup = UnoRuntime
                .queryInterface(XTextSectionsSupplier.class, xDoc);
        XNameAccess xSections = xSup.getTextSections();
        Stream.Builder<XTextSection> builder = Stream.builder();

        for (String objId : xSections.getElementNames()) {
            XTextSection xTextSection = UnoRuntime
                    .queryInterface(XTextSection.class, xSections.getByName(objId));
            builder.accept(xTextSection);
        }

        return builder.build();
    }
}
