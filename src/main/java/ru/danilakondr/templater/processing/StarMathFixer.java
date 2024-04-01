package ru.danilakondr.templater.processing;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Обработчик математических формул на языке StarMath. Необходим, чтобы
 * исправить некоторые косяки генератора MathML из LaTeX, который используется
 * в Pandoc: при обработке различного рода уравнений с переносами на следующую
 * строку получается некорректная с точки зрения LibreOffice формула, поскольку
 * с одного из концов у оператора не хватает выражения.
 * <p>
 * Язык StarMath, используемый в LibreOffice, достаточно прост,
 * поэтому были использованы регулярные выражения.
 *
 * @author Данила А. Кондратенко
 * @since 0.2.3
 */
public class StarMathFixer {
    /**
     * Регулярное выражение для обрыва выражения слева
     */
    private static final Pattern left = Pattern.compile("#\\s*([*/&|=<>]|cdot|times|div)");
    /**
     * Регулярное выражение для обрыва выражения справа
     */
    private static final Pattern right = Pattern.compile("([\\\\+\\-/&|=<>]|cdot|times|div|plusminus|minusplus)\\s*#");

    private static final int LATIN_ALPHABET_SIZE = 26;

    private static final String[] greekCapital = new String[]{
            "%ALPHA", "%BETA", "%GAMMA", "%DELTA", "%EPSILON",
            "%ZETA", "%ETA", "%THETA", "%IOTA", "%KAPPA",
            "%LAMBDA", "%MU", "%NU", "%XI", "%OMICRON",
            "%PI", "%RHO", null, "%SIGMA", "%TAU", "%UPSILON",
            "%PHI", "%CHI", "%PSI", "%OMEGA"
    };

    private static final String[] greekItalicCapital = new String[]{
            "%iALPHA", "%iBETA", "%iGAMMA", "%iDELTA", "%iEPSILON",
            "%iZETA", "%iETA", "%iTHETA", "%iIOTA", "%iKAPPA",
            "%iLAMBDA", "%iMU", "%iNU", "%iXI", "%iOMICRON",
            "%iPI", "%iRHO", null, "%iSIGMA", "%iTAU", "%iUPSILON",
            "%iPHI", "%iCHI", "%iPSI", "%iOMEGA"
    };

    private static final String[] greekSmall = new String[]{
            "%alpha", "%beta", "%gamma", "%delta", "%varepsilon",
            "%zeta", "%eta", "%theta", "%iota", "%kappa",
            "%lambda", "%mu", "%nu", "%xi", "%omicron",
            "%pi", "%rho", "%varsigma", "%sigma", "%tau", "%upsilon",
            "%varphi", "%chi", "%psi", "%omega", "partial",
            "%epsilon", "%vartheta", null, "%phi", "%varrho", "%varpi"
    };

    private static final String[] greekItalicSmall = new String[]{
            "%ialpha", "%ibeta", "%igamma", "%idelta", "%ivarepsilon",
            "%izeta", "%ieta", "%itheta", "%iiota", "%ikappa",
            "%ilambda", "%imu", "%inu", "%ixi", "%iomicron",
            "%ipi", "%irho", "%ivarsigma", "%isigma", "%itau", "%iupsilon",
            "%ivarphi", "%ichi", "%ipsi", "%iomega", "partial",
            "%iepsilon", "%ivartheta", null, "%iphi", "%ivarrho", "%ivarpi"
    };

    /**
     * Обрабочик одной формулы на языке StarMath.
     *
     * @param formula формула, которую нужно исправить
     */
    public static String fixFormula(String formula) {
        Matcher lMatch = left.matcher(formula);
        String sFormula1 = lMatch.replaceAll("# {} $1");

        Matcher rMatch = right.matcher(sFormula1);
        String sFormula2 = rMatch.replaceAll("$1 {} #");

        return fixCharacters(sFormula2);
    }

    public static String fixCharacters(String formula) {
        final int CODEPOINT_BOLD_CAPITAL_A = 119808;
        final int CODEPOINT_BOLD_SMALL_A = 119834;
        final int CODEPOINT_ITALIC_CAPITAL_A = 119860;
        final int CODEPOINT_ITALIC_SMALL_A = 119886;
        final int CODEPOINT_BOLD_ITALIC_CAPITAL_A = 119912;
        final int CODEPOINT_BOLD_ITALIC_SMALL_A = 119938;

        String f = formula;
        f = fixLatinAlphabet(f, CODEPOINT_BOLD_CAPITAL_A, CODEPOINT_BOLD_SMALL_A, "{bold nitalic %c}");
        f = fixLatinAlphabet(f, CODEPOINT_ITALIC_CAPITAL_A, CODEPOINT_ITALIC_SMALL_A, "{italic %c}");
        f = fixLatinAlphabet(f, CODEPOINT_BOLD_ITALIC_CAPITAL_A, CODEPOINT_BOLD_ITALIC_SMALL_A, "{bold italic %c}");
        f = fixGreekCapitals(f, false, false);
        f = fixGreekCapitals(f, false, true);
        f = fixGreekCapitals(f, true, false);
        f = fixGreekCapitals(f, true, true);
        f = fixGreekSmalls(f, false, false);
        f = fixGreekSmalls(f, false, true);
        f = fixGreekSmalls(f, true, false);
        f = fixGreekSmalls(f, true, true);
        return f;
    }

    private static String fixLatinAlphabet(String f, int firstCapital, int firstSmall, String replace) {
        String x = f;
        for (int i = 0; i < LATIN_ALPHABET_SIZE; i++) {
            String capital = Character.toString(firstCapital + i);
            String small = Character.toString(firstSmall + i);

            String capReplacement = String.format(replace, (char)('A' + i));
            String smallReplacement = String.format(replace, (char)('a' + i));

            x = x.replace(capital, capReplacement).replace(small, smallReplacement);
        }
        return x;
    }

    private static String fixGreekCapitals(String f, boolean italic, boolean bold) {
        String x = f;

        final int[] greekPlainCapitals = new int[]{
                0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0396, 0x0397, 0x0398,
                0x0399, 0x039A, 0x039B, 0x039C, 0x039D, 0x039E, 0x039F, 0x03A0,
                0x03A1, 0x03F4, 0x03A3, 0x03A4, 0x03A5, 0x03A6, 0x03A7, 0x03A8,
                0x03A9
        };

        for (int i = 0; i < greekCapital.length; i++) {
            if (greekCapital[i] == null)
                continue;

            String capital = "";
            if (!bold && !italic)
                capital = Character.toString(greekPlainCapitals[i]);
            else if (bold && !italic)
                capital = Character.toString(0x1D6A8 + i);
            else if (!bold && italic)
                capital = Character.toString(0x1D6E2 + i);
            else if (bold && italic)
                capital = Character.toString(0x1D71C + i);

            String capReplacement = (bold ? "bold " : "") + (!italic ? greekCapital[i] : greekItalicCapital[i]);
            x = x.replace(capital, capReplacement);
        }

        return x;
    }

    private static String fixGreekSmalls(String f, boolean italic, boolean bold) {
        String x = f;
        int k = 0;

        final int[] greekPlainSmalls = new int[]{
                0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B6, 0x03B7, 0x03B8,
                0x03B9, 0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BE, 0x03BF, 0x03C0,
                0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C5, 0x03C6, 0x03C7, 0x03C8,
                0x03C9, 0x2202, 0x03F5, 0x03D1, 0x03F0, 0x03D5, 0x03F1, 0x03D6
        };

        for (int i = 0; i < greekSmall.length; i++) {
            if (greekSmall[i] == null)
                continue;

            String small = "";
            if (!bold && !italic)
                small = Character.toString(greekPlainSmalls[i]);
            else if (bold && !italic)
                small = Character.toString(0x1D6C2 + i);
            else if (!bold && italic)
                small = Character.toString(0x1D6FC + i);
            else if (bold && italic)
                small = Character.toString(0x1D736 + i);

            String smallReplacement = (bold ? "bold " : "") + (!italic ? greekSmall[i] : greekItalicSmall[i]);
            x = x.replace(small, smallReplacement);
        }

        return x;
    }
}
