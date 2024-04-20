package ru.danilakondr.templater.app;

/**
 * Исключение при нахождении LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class LibreOfficeException extends Exception {
    public LibreOfficeException(String message) {
        super(message);
    }
}
