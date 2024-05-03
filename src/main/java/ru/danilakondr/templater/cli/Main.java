/*
 * Copyright (c) 2024 Danila A. Kondratenko
 *
 * This file is a part of UNO Templater.
 *
 * UNO Templater is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * UNO Templater is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with UNO Templater.  If not, see <https://www.gnu.org/licenses/>.
 */

package ru.danilakondr.templater.cli;

import com.sun.star.uno.XComponentContext;
import ru.danilakondr.templater.LibreOffice;
import ru.danilakondr.templater.LibreOfficeException;
import ru.danilakondr.templater.Templater;

import java.io.FileNotFoundException;
import java.io.IOException;

import ru.danilakondr.templater.BuildVersion;

public class Main {
    private static void checkHelp(String[] args) {
        for (String arg : args) {
            if (arg.equals("-h") || arg.equals("--help")) {
                printVersion();
                CommandLineArgs.Parser.printHelp();
                System.exit(0);
            }
        }
    }

    private static void checkVersion(String[] args) {
        for (String arg : args) {
            if (arg.equals("--version")) {
                printVersion();
                System.exit(0);
            }
        }
    }

    private static void printVersion() {
        System.out.printf("UNO Templater %s%n", BuildVersion.getVersion());
    }

    public static void main(String[] args) {
        checkVersion(args);
        checkHelp(args);

        CommandLineArgs templaterArgs = null;
        try {
            templaterArgs = CommandLineArgs.Parser.parseCommandLine(args);
        }
        catch (IllegalArgumentException e) {
            System.err.println(e);
            System.exit(-1);
        }

        Templater templater = new Templater();
        XComponentContext xContext = null;
        try {
            xContext = LibreOffice.bootstrap();
            templater.setContext(xContext);
        }
        catch (LibreOfficeException e) {
            System.err.println(e.getMessage());
        }
        catch (IllegalArgumentException e) {
            System.err.println("Invalid argument: " + e.getMessage());
        }
        catch (FileNotFoundException e) {
            System.err.printf("%s: file not found%n", e.getMessage());
        }
        catch (IOException e) {
            System.err.printf("%s%n", e);
        }
        catch (Exception e) {
            e.printStackTrace(System.err);
            System.exit(-1);
        }

        int status = 0;
        try {
            try {
                templater.loadMacrosFromFile(templaterArgs.getMacroFile());
            }
            catch (FileNotFoundException e) {
                System.err.printf("%s: file not found, skipping%n", e.getMessage());
            }
            catch (NullPointerException ignored) {}

            templater.loadMacrosFromMap(templaterArgs.getMacroOverrides());
            templater.setMainTextPath(templaterArgs.getMainTextPath());
            templater.setTemplatePath(templaterArgs.getTemplatePath());
            templater.setOutputPath(templaterArgs.getOutputPath());
            templater.setVerbose(templaterArgs.isShouldBeVerbose());
            templater.setShouldEmbedFonts(templaterArgs.isShouldEmbedFonts());
            templater.setShouldOverwrite(templaterArgs.isShouldOverwrite());

            templater.processDocument();
            templater.saveDocument();
            if (templaterArgs.isShouldOverwrite())
                templater.generatePDF();
        }
        catch (IllegalArgumentException e) {
            System.err.println("Invalid argument: " + e.getMessage());
        }
        catch (RuntimeException e) {
            System.err.printf("%s%n", e.getMessage());
            e.printStackTrace(System.err);
        }
        catch (FileNotFoundException e) {
            System.err.printf("%s: file not found%n", e.getMessage());
        }
        catch (Exception e) {
            e.printStackTrace(System.err);
            status = 1;
        }
        finally {
            try {
                templater.closeDocument();
            }
            catch (Exception e) {
                e.printStackTrace(System.err);
            }
            System.exit(status);
        }
    }
}
