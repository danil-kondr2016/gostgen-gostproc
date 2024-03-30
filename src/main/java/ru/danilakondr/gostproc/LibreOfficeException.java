package ru.danilakondr.gostproc;

/**
 * Исключение при нахождении LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.1
 */
public class LibreOfficeException extends Exception {
    public LibreOfficeException(String message) {
        super(message);
    }
}
