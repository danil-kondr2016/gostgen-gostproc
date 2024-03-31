package ru.danilakondr.templater.processing;

import java.io.*;
import java.util.HashMap;

/**
 * Велосипед для обработки файлов с макросами.
 *
 * @author Данила А. Кондратенко
 * @since 0.2.1
 */
public class Macros extends HashMap<String, String> {
    public Macros(File f) throws IOException {
        this(f.getPath());
    }
    public Macros(String path) throws IOException {
        super();
        loadFile(path);
    }

    /**
     * Загружает и обрабатывает файл по заданному пути.
     *
     * @param path путь к файлу
     * @throws IOException ошибка при считывании
     */
    private void loadFile(String path) throws IOException {
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
}
