package ru.danilakondr.templater.processing;

import java.util.function.BiConsumer;

public class ProgressInformer implements BiConsumer<Integer, Integer> {
    private final String progressString;
    public ProgressInformer(String progressString) {
        this.progressString = progressString;
    }

    @Override
    public void accept(Integer progress, Integer total) {
        if (total == -1)
            System.out.printf("%s (%d)...\n", progressString, progress);
        else
            System.out.printf("%s (%d/%d)...\n", progressString, progress, total);
    }
}
