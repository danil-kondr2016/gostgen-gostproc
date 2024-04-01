package ru.danilakondr.templater.processing;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Objects;

import org.apache.commons.text.lookup.StringLookup;

/**
 * Велосипед для обработки файлов с макросами.
 *
 * @author Данила А. Кондратенко
 * @since 0.2.1
 */
public class Macros extends HashMap<String, String> implements StringLookup {
    public Macros(File f) throws IOException {
        this(f.getPath());
    }
    public Macros(String path) throws IOException {
        super();
        loadFromFile(path);
    }

    public Macros() {
        super();
    }

    /**
     * Загружает и обрабатывает файл по заданному пути.
     *
     * @param path путь к файлу
     * @throws IOException ошибка при считывании
     */
    public void loadFromFile(String path) throws IOException {
        BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(path)));

        try {
            String line;

            while ((line = reader.readLine()) != null) {
                addVariable(line);
            }
        }
        finally {
            reader.close();
        }
    }

    /**
     * Добавляет переменную из одной строки.
     *
     * @param line строка
     */
    private void addVariable(String line) {
        int eqIndex = line.indexOf('=');
        if (eqIndex > 0) {
            super.put(
                    line.substring(0, eqIndex).trim(),
                    line.substring(eqIndex+1).trim()
            );
        }
    }

    @Override
    public String lookup(String s) {
        Calendar cal = Calendar.getInstance();
        Date date = cal.getTime();

        if (s.compareTo("YEAR") == 0)
            return String.format("%04d", cal.get(Calendar.YEAR));
        if (s.compareTo("DATE") == 0)
            return SimpleDateFormat.getDateInstance(DateFormat.SHORT).format(date);
        if (s.compareTo("TIME") == 0)
            return SimpleDateFormat.getTimeInstance(DateFormat.SHORT).format(date);
        if (s.compareTo("DATETIME") == 0)
            return SimpleDateFormat.getDateTimeInstance(DateFormat.SHORT, DateFormat.SHORT).format(date);

        if (s.matches("DATETIME\\((.*?)\\)"))
            return formatDateTime(s, cal);

        return get(s);
    }

    private String formatDateTime(String s, Calendar cal) {
        DateFormat fmt = new SimpleDateFormat(s.replaceAll("DATETIME\\((.*?)\\)", "$1"));
        return fmt.format(cal.getTime());
    }
}
