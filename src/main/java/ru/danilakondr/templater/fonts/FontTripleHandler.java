package ru.danilakondr.templater.fonts;


import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.OptionDef;
import org.kohsuke.args4j.spi.OptionHandler;
import org.kohsuke.args4j.spi.Parameters;
import org.kohsuke.args4j.spi.Setter;

public class FontTripleHandler extends OptionHandler<FontTriple> {
    public FontTripleHandler(CmdLineParser parser, OptionDef option, Setter<FontTriple> setter) {
        super(parser, option, setter);
    }

    @Override
    public int parseArguments(Parameters parameters) throws CmdLineException {
        String x = parameters.getParameter(0);
        String[] fonts = x.split(";");
        FontTriple triple = new FontTriple(fonts[0], fonts[1], fonts[2]);
        setter.addValue(triple);

        return 1;
    }

    @Override
    public String getDefaultMetaVariable() {
        return "TRIPLE";
    }
}
