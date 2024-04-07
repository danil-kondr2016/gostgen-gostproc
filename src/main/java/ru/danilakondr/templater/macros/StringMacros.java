package ru.danilakondr.templater.macros;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.commons.text.lookup.StringLookup;

/**
 * Обёртка над классом Properties, которую можно использовать
 * с StringSubstitutor.
 *
 * @author Данила А. Кондратенко
 * @since 0.3.2
 */
public class StringMacros implements StringLookup {
    private final Properties props;

    public StringMacros() {
        this.props = new Properties();
    }

    /**
     * Загружает строковые макросы из файла по заданному пути.
     *
     * @param path путь к файлу
     * @throws IOException ошибка при работе с файлом
     */
    public void loadFromFile(String path) throws IOException {
        InputStreamReader reader = new InputStreamReader(
                new FileInputStream(path),
                StandardCharsets.UTF_8
        );

        props.load(reader);
    }

    /**
     * Загружает строковые макросы из словаря.
     *
     * @param x словарь
     */
    public void loadFromMap(Map<String, String> x) {
        x.forEach(this.props::setProperty);
    }

    @Override
    public String lookup(String s) {
        Calendar cal = Calendar.getInstance();
        Date date = cal.getTime();

        if (s.compareTo("YEAR") == 0)
            return String.format("%04d", cal.get(Calendar.YEAR));
        if (s.compareTo("DATE") == 0)
            return SimpleDateFormat.getDateInstance(DateFormat.SHORT)
                    .format(date);
        if (s.compareTo("TIME") == 0)
            return SimpleDateFormat.getTimeInstance(DateFormat.SHORT)
                    .format(date);
        if (s.compareTo("DATETIME") == 0)
            return SimpleDateFormat.getDateTimeInstance(DateFormat.SHORT, DateFormat.SHORT)
                    .format(date);

        if (s.matches("DATETIME\\((.*?)\\)"))
            return new SimpleDateFormat(s.replaceAll("DATETIME\\((.*?)\\)", "$1"))
                    .format(date);

        return props.getProperty(s);
    }
}
