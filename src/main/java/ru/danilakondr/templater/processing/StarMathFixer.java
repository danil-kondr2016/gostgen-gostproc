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
        return f;
    }

    private static String fixLatinAlphabet(String f, int firstCapital, int firstSmall, String replace) {
        String x = f;
        for (int i = 0; i < LATIN_ALPHABET_SIZE; i++) {
            String capital = Character.toString(firstCapital + i);
            String small = Character.toString(firstSmall + i);

            String capReplacement = String.format(replace, (char) ('A' + i));
            String smallReplacement = String.format(replace, (char) ('a' + i));

            x = x.replace(capital, capReplacement).replace(small, smallReplacement);
        }
        return x;
    }
}
