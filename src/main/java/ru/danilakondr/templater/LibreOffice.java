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

package ru.danilakondr.templater;

import com.sun.star.uno.XComponentContext;
import ooo.connector.BootstrapSocketConnector;

import java.io.*;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.Properties;

import org.apache.commons.lang3.SystemUtils;

/**
 * Класс, который обеспечивает нахождение главного исполняемого файла
 * LibreOffice и запускает LibreOffice в фоновом режиме. Поддерживается
 * определение LibreOffice для трёх платформ: Windows, Linux и macOS.
 * Есть возможность указания пути к LibreOffice через переменную среды
 * <code>LIBREOFFICE_HOME</code>.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class LibreOffice {
    private static final String LIBREOFFICE_HOME_PROPERTY = "libreoffice.home";

    private static final String LIBREOFFICE_PROPERTIES_FILE = "libreoffice.properties";
    /**
     * Ищет LibreOffice.
     *
     * @return путь к директории, где хранится soffice.bin
     * @throws LibreOfficeException исключение, обозначающее ошибку
     *                              при поиске LibreOffice
     * @throws IOException исключение, обозначающее ошибку при обращении
     *                     к директории
     */
    private static String find() throws LibreOfficeException, IOException {
        String loHomePath = getLibreOfficeHomeSetting();
        if (loHomePath != null) {
            File loHomeDir = new File(loHomePath);
            if (!loHomeDir.exists())
                throw new LibreOfficeException("LibreOffice not found");
            return loHomePath;
        }

        if (SystemUtils.IS_OS_WINDOWS)
            return findLibreOfficeInWindows();
        if (SystemUtils.IS_OS_LINUX)
            return findLibreOfficeInLinux();
        if (SystemUtils.IS_OS_MAC_OSX)
            return findLibreOfficeInMacOS();

        throw new LibreOfficeException("Unrecognized platform");
    }

    private static String getLibreofficeHomeProperty(Path p) throws IOException {
        String loHomePath = null;
        Properties loProps = new Properties();

        File propsFile = p.toFile();
        if (propsFile.exists()) {
            loProps.load(new InputStreamReader(new FileInputStream(propsFile)));
        }

        loHomePath = loProps.getProperty(LIBREOFFICE_HOME_PROPERTY);

        return loHomePath;
    }

    private static String getLibreOfficeHomeSetting() throws IOException {
        String loHomePath = System.getenv("LIBREOFFICE_HOME");
        if (loHomePath == null) {
            loHomePath = System.getProperty(LIBREOFFICE_HOME_PROPERTY);
        }
        if (loHomePath == null) {
            loHomePath = getLibreofficeHomeProperty(Path.of(System.getProperty("user.home"), LIBREOFFICE_PROPERTIES_FILE));
        }
        if (loHomePath == null) {
            loHomePath = getLibreofficeHomeProperty(Path.of(".", LIBREOFFICE_PROPERTIES_FILE));
        }

        return loHomePath;
    }

    private static String findLibreOfficeInWindows() throws LibreOfficeException, IOException {
        String windir = System.getenv("windir");
        File progFiles
                = Path.of(windir, "..",
                "Program Files", "LibreOffice", "program").toFile();
        File progFiles86
                = Path.of(windir, "..",
                "Program Files (x86)", "LibreOffice", "program").toFile();

        if (progFiles.exists())
            return progFiles.getCanonicalPath();
        if (progFiles86.exists())
            return progFiles.getCanonicalPath();

        throw new LibreOfficeException("LibreOffice not found");
    }

    private static String findLibreOfficeInLinux() throws LibreOfficeException, IOException {
        File libLO = new File("/usr/lib/libreoffice/program");
        File opt = new File("/opt");

        if (libLO.exists())
            return libLO.getCanonicalPath();

        try {
            return
            Arrays.stream(Objects.requireNonNull(opt.listFiles()))
                    .filter(File::exists)
                    .filter((f) -> f.toPath().getFileName().startsWith("libreoffice"))
                    .map(File::getAbsolutePath)
                    .max(String::compareTo)
                    .orElseThrow();
        }
        catch (NullPointerException | NoSuchElementException ignored) {}

        throw new LibreOfficeException("LibreOffice not found");
    }

    private static String findLibreOfficeInMacOS() throws LibreOfficeException, IOException {
        File app = new File("/Applications/LibreOffice.cli/Contents/MacOS");
        if (!app.exists())
            throw new LibreOfficeException("LibreOffice not found");

        return app.getCanonicalPath();
    }

    /**
     * Запускает LibreOffice.
     */
    public static XComponentContext bootstrap() throws Exception {
        String path = find();
        return BootstrapSocketConnector.bootstrap(path);
    }
}

