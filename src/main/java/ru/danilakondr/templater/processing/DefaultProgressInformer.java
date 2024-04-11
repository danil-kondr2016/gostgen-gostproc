package ru.danilakondr.templater.processing;

public class DefaultProgressInformer implements ProgressInformer {
    private final String progressString;

    public DefaultProgressInformer(String progressString) {
        this.progressString = progressString;
    }

    @Override
    public void inform(int current, int total) {
        if (current == -1 && total == -1)
            System.out.println(progressString + "...");
        else if (total == -1)
            System.out.printf("%s (%d)...%n", progressString, current);
        else
            System.out.printf("%s (%d/%d)...%n", progressString, current, total);
    }
}