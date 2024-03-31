package ru.danilakondr.templater.processing;

import java.io.*;
import java.util.HashMap;

public class Definitions extends HashMap<String, String> {
    public Definitions(File f) throws IOException {
        this(f.getPath());
    }
    public Definitions(String path) throws IOException {
        super();
        loadFile(path);
    }

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
