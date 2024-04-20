package ru.danilakondr.templater.progress;

@FunctionalInterface
public interface ProgressInformer {
    void inform(int current, int total);
}
