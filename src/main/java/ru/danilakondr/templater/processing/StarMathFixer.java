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

        return sFormula2;
    }
}
