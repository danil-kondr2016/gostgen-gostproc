package ru.danilakondr.templater.app;

import com.sun.star.uno.XComponentContext;
import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.Option;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.spi.MapOptionHandler;
import ru.danilakondr.templater.LibreOffice;
import ru.danilakondr.templater.LibreOfficeException;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

public class Main {
    static class TemplaterArguments {
        @Option(name = "-t", aliases = {"--template"}, usage = "Template file")
        public String templatePath;

        @Option(name = "-m", aliases = {"--main-text"}, usage = "Main text file")
        public String mainTextPath;

        @Option(name = "-o", aliases = {"--output"}, usage = "Output file")
        public String outputPath;

        @Option(name = "-M", aliases = {"--macros"}, usage = "Macros file")
        public String macroFile;

        @Option(name = "-e", aliases = {"--embed-fonts"}, usage = "Embed fonts")
        public boolean embedFonts;

        @Option(name = "-f", aliases = {"--force", "--overwrite"}, usage = "Overwrite output file")
        public boolean overwrite;

        @Option(name = "-P", aliases = {"--pdf", "--make-pdf"}, usage = "Generate PDF file")
        public boolean generatePDF;

        @Option(name = "-D", usage = "Specify macro", handler = MapOptionHandler.class)
        public HashMap<String, String> macroOverrides;

        @Option(name = "-h", aliases = {"--help", "-?"}, help = true)
        public boolean showHelp;

        @Option(name="-v", aliases={"--verbose"})
        public boolean verbose;
    }

    public static void main(String[] args) {
        TemplaterArguments templaterArguments = new TemplaterArguments();
        CmdLineParser parser = new CmdLineParser(templaterArguments);

        try {
            parser.parseArgument(args);
        }
        catch (CmdLineException e) {
            System.err.println(e.getMessage());
            System.exit(-1);
        }

        if (templaterArguments.showHelp) {
            System.out.println("Usage:");
            System.out.print("templater ");
            parser.printSingleLineUsage(System.out);
            System.out.println();
            System.out.println();
            parser.printUsage(System.out);
            System.exit(0);
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
                templater.loadMacrosFromFile(templaterArguments.macroFile);
            }
            catch (FileNotFoundException e) {
                System.err.printf("%s: file not found, skipping%n", e.getMessage());
            }

            templater.loadMacrosFromMap(templaterArguments.macroOverrides);
            templater.setMainTextPath(templaterArguments.mainTextPath);
            templater.setTemplatePath(templaterArguments.templatePath);
            templater.setOutputPath(templaterArguments.outputPath);
            templater.setVerbose(templaterArguments.verbose);
            templater.setShouldEmbedFonts(templaterArguments.embedFonts);
            templater.setShouldOverwrite(templaterArguments.overwrite);

            templater.processDocument();
            templater.saveDocument();
            if (templaterArguments.generatePDF)
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
