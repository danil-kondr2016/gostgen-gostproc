package ru.danilakondr.templater.processing;

@FunctionalInterface
public interface ProgressInformer {
    void inform(int current, int total);
}
