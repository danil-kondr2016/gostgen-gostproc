package ru.danilakondr.templater.fonts;

import java.awt.*;
import java.util.Set;

public class FontTriple {
    public final String serif, sans, mono;
    public static final FontTriple LIBERATION = new FontTriple("Liberation Serif", "Liberation Sans", "Liberation Mono");
    public static final FontTriple MICROSOFT = new FontTriple("Times New Roman", "Arial", "Courier New");
    public static final FontTriple GOOGLE = new FontTriple("Tinos", "Arimo", "Cousine");
    public static final FontTriple PARATYPE_ASTRA = new FontTriple("PT Astra Serif", "PT Astra Sans", "PT Mono");
    public static final FontTriple PARATYPE = new FontTriple("PT Serif", "PT Sans", "PT Mono");
    private static Set<String> existingFontFamilies = null;

    public FontTriple(String serif, String sans, String mono) {
        this.serif = serif;
        this.sans = sans;
        this.mono = mono;
    }

    public static FontTriple provide() {
        if (MICROSOFT.exists())
            return MICROSOFT;
        if (PARATYPE_ASTRA.exists())
            return PARATYPE_ASTRA;
        if (PARATYPE.exists())
            return PARATYPE;
        if (LIBERATION.exists())
            return LIBERATION;
        if (GOOGLE.exists())
            return GOOGLE;

        return null;
    }

    public static FontTriple provide(FontTriple custom) {
        if (custom.exists())
            return custom;

        return provide();
    }

    public static boolean exists(FontTriple f) {
        if (existingFontFamilies == null) {
            synchronized (FontTriple.class) {
                existingFontFamilies = Set.of(GraphicsEnvironment
                        .getLocalGraphicsEnvironment()
                        .getAvailableFontFamilyNames());
            }
        }

        int hasFonts = 0;
        if (existingFontFamilies.contains(f.serif))
            hasFonts++;
        if (existingFontFamilies.contains(f.sans))
            hasFonts++;
        if (existingFontFamilies.contains(f.mono))
            hasFonts++;

        return hasFonts == 3;
    }

    public boolean exists() {
        return exists(this);
    }

    @Override
    public String toString() {
        return String.format("%s;%s;%s", this.serif, this.sans, this.mono);
    }
}
