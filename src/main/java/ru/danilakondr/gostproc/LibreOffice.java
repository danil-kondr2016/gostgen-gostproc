package ru.danilakondr.gostproc;

import com.sun.star.uno.XComponentContext;
import ooo.connector.BootstrapSocketConnector;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.function.Predicate;
import java.io.IOException;

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
        // LIBREOFFICE_HOME - папка, где находятся исполняемые файлы
        // LibreOffice.
        // На Windows это обычно C:\Program Files\LibreOffice\program.
        // На Linux это обычно /usr/lib/libreoffice/program.

        String lohome = System.getenv("LIBREOFFICE_HOME");
        if (lohome != null) {
            File fLOHome = new File(lohome);
            if (!fLOHome.exists())
                throw new LibreOfficeException("LibreOffice not found");
            return lohome;
        }

        if (SystemUtils.IS_OS_WINDOWS)
            return findLibreOfficeInWindows();
        if (SystemUtils.IS_OS_LINUX)
            return findLibreOfficeInLinux();
        if (SystemUtils.IS_OS_MAC_OSX)
            return findLibreOfficeInMacOS();

        throw new LibreOfficeException("Unrecognized platform");
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
        File app = new File("/Applications/LibreOffice.app/Contents/MacOS");
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

