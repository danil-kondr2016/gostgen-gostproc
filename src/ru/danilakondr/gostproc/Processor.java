package ru.danilakondr.gostproc;

import com.sun.star.text.XTextDocument;

public abstract class Processor {
    protected XTextDocument xDoc;

    public Processor(XTextDocument xDoc) {
        this.xDoc = xDoc;
    }

    public abstract void process() throws Exception;
}
