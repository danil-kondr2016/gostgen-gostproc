package ru.danilakondr.templater.progress;

public class DefaultProgressInformer implements ProgressInformer {
    private String progressString;
    private boolean silent = false;

    public DefaultProgressInformer(String progressString) {
        this.progressString = progressString;
    }

    public void setProgressString(String progressString) {
        this.progressString = progressString;
    }

    public void setSilent(boolean silent) {
        this.silent = silent;
    }

    @Override
    public void inform(int current, int total) {
        if (silent)
            return;

        if (current == -1 && total == -1)
            System.out.println(progressString + "...");
        else if (total == -1)
            System.out.printf("%s (%d)...%n", progressString, current);
        else
            System.out.printf("%s (%d/%d)...%n", progressString, current, total);
    }
}