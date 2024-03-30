package ru.danilakondr.gostproc;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;

/**
 * Установщик стиля первой страницы.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class FirstPageStyleSetter extends Processor {
    private XText xText;

    public FirstPageStyleSetter(XTextDocument xDoc) {
        super(xDoc);
        this.xText = xDoc.getText();
    }

    @Override
    public void process() throws Exception {
        XEnumerationAccess xEnumAccess = UnoRuntime.queryInterface(
                XEnumerationAccess.class,
                xText
        );
        XEnumeration xParEnum = xEnumAccess.createEnumeration();
        if (xParEnum.hasMoreElements()) {
            Object oFirstPar = xParEnum.nextElement();
            XPropertySet xFirstParProp = UnoRuntime.queryInterface(
                    XPropertySet.class,
                    oFirstPar
            );
            xFirstParProp.setPropertyValue("PageDescName", "First Page");
        }
    }
}
