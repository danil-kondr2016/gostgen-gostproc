package ru.danilakondr.gostproc;

import com.sun.star.text.XTextDocument;

/**
 * Базовый класс обработчиков. Содержит объект документа
 * и метод <code>process()</code>, который осуществляет обработку.
 *
 * @author Данила А. Кондратенко
 * @since 2024.03.26
 */
public abstract class Processor {
    protected XTextDocument xDoc;

    public Processor(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    /**
     * Абстрактный метод обработки. Требует реализации.
     * @throws Exception
     */
    public abstract void process() throws Exception;
}
