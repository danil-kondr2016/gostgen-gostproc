package ru.danilakondr.templater.processing;

import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.text.RelOrientation;
import com.sun.star.text.XTextDocument;
import com.sun.star.container.*;
import com.sun.star.text.XTextGraphicObjectsSupplier;
import com.sun.star.uno.UnoRuntime;

import java.util.concurrent.atomic.AtomicInteger;

public class ImageWidthFixer extends Processor {
    private final XNameAccess graphicObjects;

    public ImageWidthFixer(XTextDocument xDoc) {
        super(xDoc);

        graphicObjects = UnoRuntime
                .queryInterface(XTextGraphicObjectsSupplier.class, xDoc)
                .getGraphicObjects();
    }

    @Override
    public void process() throws Exception {
        System.out.println("Resizing image objects...");

        String[] names = graphicObjects.getElementNames();

        AtomicInteger i = new AtomicInteger(0);
        for (String objId : names) {
            System.out.printf("Processing image object %d/%d...\n", i.incrementAndGet(), names.length);
            XPropertySet xObject = UnoRuntime
                    .queryInterface(XPropertySet.class, graphicObjects.getByName(objId));

            Size actualSize = (Size) xObject.getPropertyValue("ActualSize");

            if (actualSize.Width > 16500) {
                long height = 16500L * actualSize.Height / actualSize.Width;

                if (height <= 27700L) {
                    xObject.setPropertyValue("Width", 16500);
                    xObject.setPropertyValue("Height", Long.valueOf(height).intValue());
                }
                else {
                    long width = 27700L * actualSize.Width / actualSize.Height;

                    xObject.setPropertyValue("Width", Long.valueOf(width).intValue());
                    xObject.setPropertyValue("Height", 27700);
                }
            }
        }
    }
}
