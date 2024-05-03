package ru.danilakondr.templater.app;

import java.util.HashMap;
import java.util.Properties;

import org.apache.commons.cli.*;

public class CommandLineArgs {
    private String templatePath;
    private String mainTextPath;
    private String outputPath;
    private String macroFile;
    private boolean shouldEmbedFonts;
    private boolean shouldOverwrite;
    private boolean shouldGeneratePDF;
    private final Properties macroOverrides;
    private boolean shouldBeVerbose;

    public CommandLineArgs() {
        this.macroOverrides = new Properties();
    }

    public String getTemplatePath() {
        return templatePath;
    }

    public String getMainTextPath() {
        return mainTextPath;
    }

    public String getOutputPath() {
        return outputPath;
    }

    public Properties getMacroOverrides() {
        return macroOverrides;
    }

    public String getMacroFile() {
        return macroFile;
    }

    public boolean isShouldBeVerbose() {
        return shouldBeVerbose;
    }

    public boolean isShouldEmbedFonts() {
        return shouldEmbedFonts;
    }

    public boolean isShouldGeneratePDF() {
        return shouldGeneratePDF;
    }

    public boolean isShouldOverwrite() {
        return shouldOverwrite;
    }

    public void setMacroFile(String macroFile) {
        this.macroFile = macroFile;
    }

    public void setMainTextPath(String mainTextPath) {
        this.mainTextPath = mainTextPath;
    }

    public void setOutputPath(String outputPath) {
        this.outputPath = outputPath;
    }

    public void setShouldBeVerbose(boolean shouldBeVerbose) {
        this.shouldBeVerbose = shouldBeVerbose;
    }

    public void setShouldEmbedFonts(boolean shouldEmbedFonts) {
        this.shouldEmbedFonts = shouldEmbedFonts;
    }

    public void setShouldGeneratePDF(boolean shouldGeneratePDF) {
        this.shouldGeneratePDF = shouldGeneratePDF;
    }

    public void setShouldOverwrite(boolean shouldOverwrite) {
        this.shouldOverwrite = shouldOverwrite;
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

    public void addMacroOverride(String key, String value) {
        macroOverrides.put(key, value);
    }

    public static class Parser {
        private static final Option OPTION_TEMPLATE;
        private static final Option OPTION_MAIN_TEXT;
        private static final Option OPTION_OUTPUT;
        private static final Option OPTION_MACRO_FILE;
        private static final Option OPTION_EMBED_FONTS;
        private static final Option OPTION_OVERWRITE;
        private static final Option OPTION_GENERATE_PDF;
        private static final Option OPTION_MACRO_DEF;
        private static final Option OPTION_VERBOSE;
        private static final Option OPTION_SHOW_HELP;
        private static final Options opts;

        static {
            OPTION_TEMPLATE = Option.builder("t")
                    .longOpt("template")
                    .argName("TEMPLATE")
                    .desc("Template file")
                    .hasArg()
                    .required()
                    .build();

            OPTION_MAIN_TEXT = Option.builder("m")
                    .longOpt("main")
                    .argName("MAINFILE")
                    .desc("Main text file")
                    .hasArg()
                    .required()
                    .build();

            OPTION_OUTPUT = Option.builder("o")
                    .longOpt("output")
                    .argName("OUTFILE")
                    .desc("Output file")
                    .hasArg()
                    .required()
                    .build();

            OPTION_MACRO_FILE = Option.builder("M")
                    .longOpt("macros")
                    .argName("MACRO_FILE")
                    .desc("File with string macros")
                    .hasArg()
                    .build();

            OPTION_EMBED_FONTS = Option.builder("e")
                    .longOpt("embed-fonts")
                    .desc("Embed fonts")
                    .build();

            OPTION_OVERWRITE = Option.builder("f")
                    .longOpt("force")
                    .desc("Force overwrite")
                    .build();

            OPTION_GENERATE_PDF = Option.builder("P")
                    .longOpt("make-pdf")
                    .desc("Generate PDF file")
                    .build();

            OPTION_MACRO_DEF = Option.builder("D")
                    .hasArgs()
                    .valueSeparator()
                    .build();

            OPTION_VERBOSE = Option.builder("v")
                    .longOpt("verbose")
                    .desc("Print process messages")
                    .build();

            OPTION_SHOW_HELP = Option.builder("h")
                    .longOpt("help")
                    .desc("Print help message")
                    .build();

            opts = new Options()
                    .addOption(OPTION_TEMPLATE)
                    .addOption(OPTION_MAIN_TEXT)
                    .addOption(OPTION_OUTPUT)
                    .addOption(OPTION_MACRO_FILE)
                    .addOption(OPTION_EMBED_FONTS)
                    .addOption(OPTION_OVERWRITE)
                    .addOption(OPTION_GENERATE_PDF)
                    .addOption(OPTION_MACRO_DEF)
                    .addOption(OPTION_VERBOSE)
                    .addOption(OPTION_SHOW_HELP)
                    ;
        }

        public static CommandLineArgs parseCommandLine(String[] cmdLine) {
            try {
                CommandLineParser parser = new DefaultParser();
                CommandLine cmd = parser.parse(opts, cmdLine);
                CommandLineArgs result = new CommandLineArgs();

                result.setTemplatePath(cmd.getOptionValue(OPTION_TEMPLATE));
                result.setMainTextPath(cmd.getOptionValue(OPTION_MAIN_TEXT));
                result.setOutputPath(cmd.getOptionValue(OPTION_OUTPUT));
                if (cmd.hasOption(OPTION_MACRO_FILE))
                    result.setMacroFile(cmd.getOptionValue(OPTION_MACRO_FILE));
                result.setShouldOverwrite(cmd.hasOption(OPTION_OVERWRITE));
                result.setShouldBeVerbose(cmd.hasOption(OPTION_VERBOSE));
                result.setShouldGeneratePDF(cmd.hasOption(OPTION_GENERATE_PDF));
                result.setShouldEmbedFonts(cmd.hasOption(OPTION_EMBED_FONTS));

                if (cmd.hasOption(OPTION_MACRO_DEF)) {
                    Properties p = cmd.getOptionProperties(OPTION_MACRO_DEF);
                    result.getMacroOverrides().putAll(p);
                }

                return result;
            }
            catch (ParseException e) {
                throw new IllegalArgumentException(e);
            }
        }

        public static void printHelp() {
            HelpFormatter formatter = new HelpFormatter();
            formatter.printHelp("templater", opts, true);
        }
    }
}
