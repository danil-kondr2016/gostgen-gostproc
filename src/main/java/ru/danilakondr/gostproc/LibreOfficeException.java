package ru.danilakondr.gostproc;

/**
 * Исключение при нахождении LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 2024.03.29
 */
public class LibreOfficeException extends Exception {
    public LibreOfficeException(String message) {
        super(message);
    }
}
