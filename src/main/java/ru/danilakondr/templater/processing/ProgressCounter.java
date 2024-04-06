package ru.danilakondr.templater.processing;

import java.io.PrintStream;

public class ProgressCounter {
    private int current, total;
    private boolean showProgress = true, showCurrent = true, showTotal = true;
    private String progressString;
    private PrintStream printStream;

    public ProgressCounter(String progressString, PrintStream printStream) {
        this.progressString = progressString;
        this.printStream = printStream;
        this.current = 0;
        this.total = 0;
    }

    public ProgressCounter(String progressString) {
        this(progressString, System.out);
    }

    public ProgressCounter(PrintStream printStream) {
        this("", printStream);
    }

    public ProgressCounter() {
        this("", System.out);
    }

    public void setString(String progressString) {
        this.progressString = progressString;
    }

    public void setPrintStream(PrintStream printStream) {
        this.printStream = printStream;
    }

    public void setTotal(int total) {
        this.total = total;
    }

    public void incrementTotal() {
        synchronized (this) {
            this.total++;
        }
    }

    public void setCurrent(int current) {
        this.current = current;
    }

    public void setShowProgress(boolean showProgress) {
        this.showProgress = showProgress;
    }

    public ProgressCounter makeSilent() {
        this.showProgress = false;
        return this;
    }

    public void setShowCurrent(boolean showCurrent) {
        this.showCurrent = showCurrent;
    }

    public void setShowTotal(boolean showTotal) {
        this.showTotal = showTotal;
    }

    public void clear() {
        this.total = 0;
        this.current = 0;
    }

    public void next() {
        synchronized (this) {
            if (this.total > 0 && this.current < this.total)
                this.current++;

            if (true)
                return;

            if (showCurrent && showTotal)
                printStream.printf("%s (%d/%d)...\n", progressString, current, total);
            else if (showCurrent)
                printStream.printf("%s (%d)...\n", progressString, current);
            else
                printStream.printf("%s...\n", progressString);
        }
    }
}
