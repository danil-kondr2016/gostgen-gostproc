package ru.danilakondr.templater.processing;

import org.apache.commons.text.StringEscapeUtils;

import java.util.Arrays;
import java.util.Map;
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
    private static final Pattern left = Pattern.compile("([#{])\\s*([*/&|=<>]|cdot|times|div)");
    /**
     * Регулярное выражение для обрыва выражения справа
     */
    private static final Pattern right = Pattern.compile("([\\\\+\\-/&|=<>]|cdot|times|div|plusminus|minusplus)\\s*([#}])");

    private static class ReplacePair {
        public final Pattern a;
        public final String b;


        private ReplacePair(Pattern a, String b) {
            this.a = a;
            this.b = b;
        }
    }

    private static final ReplacePair[] accentsReplacement = new ReplacePair[]{
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0300"), "\\{grave $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0301"), "\\{acute $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0302"), "\\{hat $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0303"), "\\{tilde $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0304"), "\\{bar $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0305"), "\\{widebar $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0306"), "\\{breve $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0307"), "\\{dot $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0308"), "\\{ddot $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u030A"), "\\{circle $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u030C"), "\\{check $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u035E"), "\\{overline $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u035F"), "\\{underline $1\\}"),
            new ReplacePair(Pattern.compile("(\\s*.{1,2}\\s*)csup\\s*\\u0360"), "\\{widetilde $1\\}"),
    };

    private static final ReplacePair[] bracesReplacement = new ReplacePair[] {
            new ReplacePair(Pattern.compile("left\\s+mline"), "left lline"),
            new ReplacePair(Pattern.compile("right\\s+mline"), "right rline"),
    };

    /**
     * Обрабочик одной формулы на языке StarMath.
     *
     * @param formula формула, которую нужно исправить
     */
    public static String fixFormula(String formula) {
        String f = formula;

        f = fixAccents(f);
        f = fixBraces(f);
        f = fixCharacters(f);
        f = fixOperators(f);

        return f;
    }

    private static String fixOperators(String formula) {
        String f = formula;

        f = left.matcher(f).replaceAll("$1 {} $2");
        f = right.matcher(f).replaceAll("$1 {} $2");

        return f;
    }

    private static String fixCharacters(String formula) {
        StringBuilder f = new StringBuilder();
        Pattern symbol = Pattern.compile("([\\u0370-\\u03FF]|[\\u2200-\\u22FF]|[\\x{1D400}-\\x{1D7FF}])");
        Matcher m = symbol.matcher(formula);

        while (m.find()) {
            if (m.group(1) != null) {
                String a = m.group(1);
                String b = symbolReplacement.getOrDefault(a, "$0");

                m.appendReplacement(f, b);
            }
            else {
                m.appendReplacement(f, "$0");
            }
        }
        m.appendTail(f);

        return f.toString();
    }

    private static String fixBraces(String f) {
        String x = f;

        for (ReplacePair p : bracesReplacement) {
            x = p.a.matcher(x).replaceAll(p.b);
        }
        return x;
    }

    private static String fixAccents(String f) {
        String x = f;

        for (ReplacePair p : accentsReplacement) {
            x = p.a.matcher(x).replaceAll(p.b);
        }

        return x;
    }

    private static final Map<String, String> symbolReplacement = Map.<String, String>ofEntries(
            Map.entry(Character.toString(913), "%ALPHA"),
            Map.entry(Character.toString(914), "%BETA"),
            Map.entry(Character.toString(915), "%GAMMA"),
            Map.entry(Character.toString(916), "%DELTA"),
            Map.entry(Character.toString(0x2206), "%DELTA"),
            Map.entry(Character.toString(917), "%EPSILON"),
            Map.entry(Character.toString(918), "%ZETA"),
            Map.entry(Character.toString(919), "%ETA"),
            Map.entry(Character.toString(920), "%THETA"),
            Map.entry(Character.toString(921), "%IOTA"),
            Map.entry(Character.toString(922), "%KAPPA"),
            Map.entry(Character.toString(923), "%LAMBDA"),
            Map.entry(Character.toString(924), "%MU"),
            Map.entry(Character.toString(925), "%NU"),
            Map.entry(Character.toString(926), "%XI"),
            Map.entry(Character.toString(927), "%OMICRON"),
            Map.entry(Character.toString(928), "%PI"),
            Map.entry(Character.toString(929), "%RHO"),
            Map.entry(Character.toString(1012), "$1"),
            Map.entry(Character.toString(931), "%SIGMA"),
            Map.entry(Character.toString(932), "%TAU"),
            Map.entry(Character.toString(933), "%UPSILON"),
            Map.entry(Character.toString(934), "%PHI"),
            Map.entry(Character.toString(935), "%CHI"),
            Map.entry(Character.toString(936), "%PSI"),
            Map.entry(Character.toString(937), "%OMEGA"),
            Map.entry(Character.toString(8711), "nabla"),
            Map.entry(Character.toString(945), "%alpha"),
            Map.entry(Character.toString(946), "%beta"),
            Map.entry(Character.toString(947), "%gamma"),
            Map.entry(Character.toString(948), "%delta"),
            Map.entry(Character.toString(949), "%varepsilon"),
            Map.entry(Character.toString(950), "%zeta"),
            Map.entry(Character.toString(951), "%eta"),
            Map.entry(Character.toString(952), "%theta"),
            Map.entry(Character.toString(953), "%iota"),
            Map.entry(Character.toString(954), "%kappa"),
            Map.entry(Character.toString(955), "%lambda"),
            Map.entry(Character.toString(956), "%mu"),
            Map.entry(Character.toString(957), "%nu"),
            Map.entry(Character.toString(958), "%xi"),
            Map.entry(Character.toString(959), "%omicron"),
            Map.entry(Character.toString(960), "%pi"),
            Map.entry(Character.toString(961), "%rho"),
            Map.entry(Character.toString(962), "%varsigma"),
            Map.entry(Character.toString(963), "%sigma"),
            Map.entry(Character.toString(964), "%tau"),
            Map.entry(Character.toString(965), "%upsilon"),
            Map.entry(Character.toString(966), "%varphi"),
            Map.entry(Character.toString(967), "%chi"),
            Map.entry(Character.toString(968), "%psi"),
            Map.entry(Character.toString(969), "%omega"),
            Map.entry(Character.toString(8706), "partial"),
            Map.entry(Character.toString(1013), "%epsilon"),
            Map.entry(Character.toString(977), "%vartheta"),
            Map.entry(Character.toString(1008), "$1"),
            Map.entry(Character.toString(981), "%phi"),
            Map.entry(Character.toString(1009), "%varrho"),
            Map.entry(Character.toString(982), "%varpi"),
            Map.entry(Character.toString(120488), "{bold %ALPHA}"),
            Map.entry(Character.toString(120489), "{bold %BETA}"),
            Map.entry(Character.toString(120490), "{bold %GAMMA}"),
            Map.entry(Character.toString(120491), "{bold %DELTA}"),
            Map.entry(Character.toString(120492), "{bold %EPSILON}"),
            Map.entry(Character.toString(120493), "{bold %ZETA}"),
            Map.entry(Character.toString(120494), "{bold %ETA}"),
            Map.entry(Character.toString(120495), "{bold %THETA}"),
            Map.entry(Character.toString(120496), "{bold %IOTA}"),
            Map.entry(Character.toString(120497), "{bold %KAPPA}"),
            Map.entry(Character.toString(120498), "{bold %LAMBDA}"),
            Map.entry(Character.toString(120499), "{bold %MU}"),
            Map.entry(Character.toString(120500), "{bold %NU}"),
            Map.entry(Character.toString(120501), "{bold %XI}"),
            Map.entry(Character.toString(120502), "{bold %OMICRON}"),
            Map.entry(Character.toString(120503), "{bold %PI}"),
            Map.entry(Character.toString(120504), "{bold %RHO}"),
            Map.entry(Character.toString(120505), "$1"),
            Map.entry(Character.toString(120506), "{bold %SIGMA}"),
            Map.entry(Character.toString(120507), "{bold %TAU}"),
            Map.entry(Character.toString(120508), "{bold %UPSILON}"),
            Map.entry(Character.toString(120509), "{bold %PHI}"),
            Map.entry(Character.toString(120510), "{bold %CHI}"),
            Map.entry(Character.toString(120511), "{bold %PSI}"),
            Map.entry(Character.toString(120512), "{bold %OMEGA}"),
            Map.entry(Character.toString(120513), "{bold nabla}"),
            Map.entry(Character.toString(120514), "{bold %alpha}"),
            Map.entry(Character.toString(120515), "{bold %beta}"),
            Map.entry(Character.toString(120516), "{bold %gamma}"),
            Map.entry(Character.toString(120517), "{bold %delta}"),
            Map.entry(Character.toString(120518), "{bold %varepsilon}"),
            Map.entry(Character.toString(120519), "{bold %zeta}"),
            Map.entry(Character.toString(120520), "{bold %eta}"),
            Map.entry(Character.toString(120521), "{bold %theta}"),
            Map.entry(Character.toString(120522), "{bold %iota}"),
            Map.entry(Character.toString(120523), "{bold %kappa}"),
            Map.entry(Character.toString(120524), "{bold %lambda}"),
            Map.entry(Character.toString(120525), "{bold %mu}"),
            Map.entry(Character.toString(120526), "{bold %nu}"),
            Map.entry(Character.toString(120527), "{bold %xi}"),
            Map.entry(Character.toString(120528), "{bold %omicron}"),
            Map.entry(Character.toString(120529), "{bold %pi}"),
            Map.entry(Character.toString(120530), "{bold %rho}"),
            Map.entry(Character.toString(120531), "{bold %varsigma}"),
            Map.entry(Character.toString(120532), "{bold %sigma}"),
            Map.entry(Character.toString(120533), "{bold %tau}"),
            Map.entry(Character.toString(120534), "{bold %upsilon}"),
            Map.entry(Character.toString(120535), "{bold %varphi}"),
            Map.entry(Character.toString(120536), "{bold %chi}"),
            Map.entry(Character.toString(120537), "{bold %psi}"),
            Map.entry(Character.toString(120538), "{bold %omega}"),
            Map.entry(Character.toString(120539), "{bold partial}"),
            Map.entry(Character.toString(120540), "{bold %epsilon}"),
            Map.entry(Character.toString(120541), "{bold %vartheta}"),
            Map.entry(Character.toString(120542), "$1"),
            Map.entry(Character.toString(120543), "{bold %phi}"),
            Map.entry(Character.toString(120544), "{bold %varrho}"),
            Map.entry(Character.toString(120545), "{bold %varpi}"),
            Map.entry(Character.toString(120546), "%iALPHA"),
            Map.entry(Character.toString(120547), "%iBETA"),
            Map.entry(Character.toString(120548), "%iGAMMA"),
            Map.entry(Character.toString(120549), "%iDELTA"),
            Map.entry(Character.toString(120550), "%iEPSILON"),
            Map.entry(Character.toString(120551), "%iZETA"),
            Map.entry(Character.toString(120552), "%iETA"),
            Map.entry(Character.toString(120553), "%iTHETA"),
            Map.entry(Character.toString(120554), "%iIOTA"),
            Map.entry(Character.toString(120555), "%iKAPPA"),
            Map.entry(Character.toString(120556), "%iLAMBDA"),
            Map.entry(Character.toString(120557), "%iMU"),
            Map.entry(Character.toString(120558), "%iNU"),
            Map.entry(Character.toString(120559), "%iXI"),
            Map.entry(Character.toString(120560), "%iOMICRON"),
            Map.entry(Character.toString(120561), "%iPI"),
            Map.entry(Character.toString(120562), "%iRHO"),
            Map.entry(Character.toString(120563), "$1"),
            Map.entry(Character.toString(120564), "%iSIGMA"),
            Map.entry(Character.toString(120565), "%iTAU"),
            Map.entry(Character.toString(120566), "%iUPSILON"),
            Map.entry(Character.toString(120567), "%iPHI"),
            Map.entry(Character.toString(120568), "%iCHI"),
            Map.entry(Character.toString(120569), "%iPSI"),
            Map.entry(Character.toString(120570), "%iOMEGA"),
            Map.entry(Character.toString(120571), "{italic nabla}"),
            Map.entry(Character.toString(120572), "%ialpha"),
            Map.entry(Character.toString(120573), "%ibeta"),
            Map.entry(Character.toString(120574), "%igamma"),
            Map.entry(Character.toString(120575), "%idelta"),
            Map.entry(Character.toString(120576), "%ivarepsilon"),
            Map.entry(Character.toString(120577), "%izeta"),
            Map.entry(Character.toString(120578), "%ieta"),
            Map.entry(Character.toString(120579), "%itheta"),
            Map.entry(Character.toString(120580), "%iiota"),
            Map.entry(Character.toString(120581), "%ikappa"),
            Map.entry(Character.toString(120582), "%ilambda"),
            Map.entry(Character.toString(120583), "%imu"),
            Map.entry(Character.toString(120584), "%inu"),
            Map.entry(Character.toString(120585), "%ixi"),
            Map.entry(Character.toString(120586), "%iomicron"),
            Map.entry(Character.toString(120587), "%ipi"),
            Map.entry(Character.toString(120588), "%irho"),
            Map.entry(Character.toString(120589), "%ivarsigma"),
            Map.entry(Character.toString(120590), "%isigma"),
            Map.entry(Character.toString(120591), "%itau"),
            Map.entry(Character.toString(120592), "%iupsilon"),
            Map.entry(Character.toString(120593), "%ivarphi"),
            Map.entry(Character.toString(120594), "%ichi"),
            Map.entry(Character.toString(120595), "%ipsi"),
            Map.entry(Character.toString(120596), "%iomega"),
            Map.entry(Character.toString(120597), "partial"),
            Map.entry(Character.toString(120598), "%iepsilon"),
            Map.entry(Character.toString(120599), "%ivartheta"),
            Map.entry(Character.toString(120600), "$1"),
            Map.entry(Character.toString(120601), "%iphi"),
            Map.entry(Character.toString(120602), "%ivarrho"),
            Map.entry(Character.toString(120603), "%ivarpi"),
            Map.entry(Character.toString(120604), "{bold %iALPHA}"),
            Map.entry(Character.toString(120605), "{bold %iBETA}"),
            Map.entry(Character.toString(120606), "{bold %iGAMMA}"),
            Map.entry(Character.toString(120607), "{bold %iDELTA}"),
            Map.entry(Character.toString(120608), "{bold %iEPSILON}"),
            Map.entry(Character.toString(120609), "{bold %iZETA}"),
            Map.entry(Character.toString(120610), "{bold %iETA}"),
            Map.entry(Character.toString(120611), "{bold %iTHETA}"),
            Map.entry(Character.toString(120612), "{bold %iIOTA}"),
            Map.entry(Character.toString(120613), "{bold %iKAPPA}"),
            Map.entry(Character.toString(120614), "{bold %iLAMBDA}"),
            Map.entry(Character.toString(120615), "{bold %iMU}"),
            Map.entry(Character.toString(120616), "{bold %iNU}"),
            Map.entry(Character.toString(120617), "{bold %iXI}"),
            Map.entry(Character.toString(120618), "{bold %iOMICRON}"),
            Map.entry(Character.toString(120619), "{bold %iPI}"),
            Map.entry(Character.toString(120620), "{bold %iRHO}"),
            Map.entry(Character.toString(120621), "$1"),
            Map.entry(Character.toString(120622), "{bold %iSIGMA}"),
            Map.entry(Character.toString(120623), "{bold %iTAU}"),
            Map.entry(Character.toString(120624), "{bold %iUPSILON}"),
            Map.entry(Character.toString(120625), "{bold %iPHI}"),
            Map.entry(Character.toString(120626), "{bold %iCHI}"),
            Map.entry(Character.toString(120627), "{bold %iPSI}"),
            Map.entry(Character.toString(120628), "{bold %iOMEGA}"),
            Map.entry(Character.toString(120629), "{bold italic nabla}"),
            Map.entry(Character.toString(120630), "{bold %ialpha}"),
            Map.entry(Character.toString(120631), "{bold %ibeta}"),
            Map.entry(Character.toString(120632), "{bold %igamma}"),
            Map.entry(Character.toString(120633), "{bold %idelta}"),
            Map.entry(Character.toString(120634), "{bold %ivarepsilon}"),
            Map.entry(Character.toString(120635), "{bold %izeta}"),
            Map.entry(Character.toString(120636), "{bold %ieta}"),
            Map.entry(Character.toString(120637), "{bold %itheta}"),
            Map.entry(Character.toString(120638), "{bold %iiota}"),
            Map.entry(Character.toString(120639), "{bold %ikappa}"),
            Map.entry(Character.toString(120640), "{bold %ilambda}"),
            Map.entry(Character.toString(120641), "{bold %imu}"),
            Map.entry(Character.toString(120642), "{bold %inu}"),
            Map.entry(Character.toString(120643), "{bold %ixi}"),
            Map.entry(Character.toString(120644), "{bold %iomicron}"),
            Map.entry(Character.toString(120645), "{bold %ipi}"),
            Map.entry(Character.toString(120646), "{bold %irho}"),
            Map.entry(Character.toString(120647), "{bold %ivarsigma}"),
            Map.entry(Character.toString(120648), "{bold %isigma}"),
            Map.entry(Character.toString(120649), "{bold %itau}"),
            Map.entry(Character.toString(120650), "{bold %iupsilon}"),
            Map.entry(Character.toString(120651), "{bold %ivarphi}"),
            Map.entry(Character.toString(120652), "{bold %ichi}"),
            Map.entry(Character.toString(120653), "{bold %ipsi}"),
            Map.entry(Character.toString(120654), "{bold %iomega}"),
            Map.entry(Character.toString(120655), "{bold partial}"),
            Map.entry(Character.toString(120656), "{bold %iepsilon}"),
            Map.entry(Character.toString(120657), "{bold %ivartheta}"),
            Map.entry(Character.toString(120658), "$1"),
            Map.entry(Character.toString(120659), "{bold %iphi}"),
            Map.entry(Character.toString(120660), "{bold %ivarrho}"),
            Map.entry(Character.toString(120661), "{bold %ivarpi}"),
            Map.entry(Character.toString(119808), "{bold A}"),
            Map.entry(Character.toString(119809), "{bold B}"),
            Map.entry(Character.toString(119810), "{bold C}"),
            Map.entry(Character.toString(119811), "{bold D}"),
            Map.entry(Character.toString(119812), "{bold E}"),
            Map.entry(Character.toString(119813), "{bold F}"),
            Map.entry(Character.toString(119814), "{bold G}"),
            Map.entry(Character.toString(119815), "{bold H}"),
            Map.entry(Character.toString(119816), "{bold I}"),
            Map.entry(Character.toString(119817), "{bold J}"),
            Map.entry(Character.toString(119818), "{bold K}"),
            Map.entry(Character.toString(119819), "{bold L}"),
            Map.entry(Character.toString(119820), "{bold M}"),
            Map.entry(Character.toString(119821), "{bold N}"),
            Map.entry(Character.toString(119822), "{bold O}"),
            Map.entry(Character.toString(119823), "{bold P}"),
            Map.entry(Character.toString(119824), "{bold Q}"),
            Map.entry(Character.toString(119825), "{bold R}"),
            Map.entry(Character.toString(119826), "{bold S}"),
            Map.entry(Character.toString(119827), "{bold T}"),
            Map.entry(Character.toString(119828), "{bold U}"),
            Map.entry(Character.toString(119829), "{bold V}"),
            Map.entry(Character.toString(119830), "{bold W}"),
            Map.entry(Character.toString(119831), "{bold X}"),
            Map.entry(Character.toString(119832), "{bold Y}"),
            Map.entry(Character.toString(119833), "{bold Z}"),
            Map.entry(Character.toString(119834), "{bold a}"),
            Map.entry(Character.toString(119835), "{bold b}"),
            Map.entry(Character.toString(119836), "{bold c}"),
            Map.entry(Character.toString(119837), "{bold d}"),
            Map.entry(Character.toString(119838), "{bold e}"),
            Map.entry(Character.toString(119839), "{bold f}"),
            Map.entry(Character.toString(119840), "{bold g}"),
            Map.entry(Character.toString(119841), "{bold h}"),
            Map.entry(Character.toString(119842), "{bold i}"),
            Map.entry(Character.toString(119843), "{bold j}"),
            Map.entry(Character.toString(119844), "{bold k}"),
            Map.entry(Character.toString(119845), "{bold l}"),
            Map.entry(Character.toString(119846), "{bold m}"),
            Map.entry(Character.toString(119847), "{bold n}"),
            Map.entry(Character.toString(119848), "{bold o}"),
            Map.entry(Character.toString(119849), "{bold p}"),
            Map.entry(Character.toString(119850), "{bold q}"),
            Map.entry(Character.toString(119851), "{bold r}"),
            Map.entry(Character.toString(119852), "{bold s}"),
            Map.entry(Character.toString(119853), "{bold t}"),
            Map.entry(Character.toString(119854), "{bold u}"),
            Map.entry(Character.toString(119855), "{bold v}"),
            Map.entry(Character.toString(119856), "{bold w}"),
            Map.entry(Character.toString(119857), "{bold x}"),
            Map.entry(Character.toString(119858), "{bold y}"),
            Map.entry(Character.toString(119859), "{bold z}"),
            Map.entry(Character.toString(119860), "{italic A}"),
            Map.entry(Character.toString(119861), "{italic B}"),
            Map.entry(Character.toString(119862), "{italic C}"),
            Map.entry(Character.toString(119863), "{italic D}"),
            Map.entry(Character.toString(119864), "{italic E}"),
            Map.entry(Character.toString(119865), "{italic F}"),
            Map.entry(Character.toString(119866), "{italic G}"),
            Map.entry(Character.toString(119867), "{italic H}"),
            Map.entry(Character.toString(119868), "{italic I}"),
            Map.entry(Character.toString(119869), "{italic J}"),
            Map.entry(Character.toString(119870), "{italic K}"),
            Map.entry(Character.toString(119871), "{italic L}"),
            Map.entry(Character.toString(119872), "{italic M}"),
            Map.entry(Character.toString(119873), "{italic N}"),
            Map.entry(Character.toString(119874), "{italic O}"),
            Map.entry(Character.toString(119875), "{italic P}"),
            Map.entry(Character.toString(119876), "{italic Q}"),
            Map.entry(Character.toString(119877), "{italic R}"),
            Map.entry(Character.toString(119878), "{italic S}"),
            Map.entry(Character.toString(119879), "{italic T}"),
            Map.entry(Character.toString(119880), "{italic U}"),
            Map.entry(Character.toString(119881), "{italic V}"),
            Map.entry(Character.toString(119882), "{italic W}"),
            Map.entry(Character.toString(119883), "{italic X}"),
            Map.entry(Character.toString(119884), "{italic Y}"),
            Map.entry(Character.toString(119885), "{italic Z}"),
            Map.entry(Character.toString(119886), "{italic a}"),
            Map.entry(Character.toString(119887), "{italic b}"),
            Map.entry(Character.toString(119888), "{italic c}"),
            Map.entry(Character.toString(119889), "{italic d}"),
            Map.entry(Character.toString(119890), "{italic e}"),
            Map.entry(Character.toString(119891), "{italic f}"),
            Map.entry(Character.toString(119892), "{italic g}"),
            Map.entry(Character.toString(119893), "{italic h}"),
            Map.entry(Character.toString(119894), "{italic i}"),
            Map.entry(Character.toString(119895), "{italic j}"),
            Map.entry(Character.toString(119896), "{italic k}"),
            Map.entry(Character.toString(119897), "{italic l}"),
            Map.entry(Character.toString(119898), "{italic m}"),
            Map.entry(Character.toString(119899), "{italic n}"),
            Map.entry(Character.toString(119900), "{italic o}"),
            Map.entry(Character.toString(119901), "{italic p}"),
            Map.entry(Character.toString(119902), "{italic q}"),
            Map.entry(Character.toString(119903), "{italic r}"),
            Map.entry(Character.toString(119904), "{italic s}"),
            Map.entry(Character.toString(119905), "{italic t}"),
            Map.entry(Character.toString(119906), "{italic u}"),
            Map.entry(Character.toString(119907), "{italic v}"),
            Map.entry(Character.toString(119908), "{italic w}"),
            Map.entry(Character.toString(119909), "{italic x}"),
            Map.entry(Character.toString(119910), "{italic y}"),
            Map.entry(Character.toString(119911), "{italic z}"),
            Map.entry(Character.toString(119912), "{bold italic A}"),
            Map.entry(Character.toString(119913), "{bold italic B}"),
            Map.entry(Character.toString(119914), "{bold italic C}"),
            Map.entry(Character.toString(119915), "{bold italic D}"),
            Map.entry(Character.toString(119916), "{bold italic E}"),
            Map.entry(Character.toString(119917), "{bold italic F}"),
            Map.entry(Character.toString(119918), "{bold italic G}"),
            Map.entry(Character.toString(119919), "{bold italic H}"),
            Map.entry(Character.toString(119920), "{bold italic I}"),
            Map.entry(Character.toString(119921), "{bold italic J}"),
            Map.entry(Character.toString(119922), "{bold italic K}"),
            Map.entry(Character.toString(119923), "{bold italic L}"),
            Map.entry(Character.toString(119924), "{bold italic M}"),
            Map.entry(Character.toString(119925), "{bold italic N}"),
            Map.entry(Character.toString(119926), "{bold italic O}"),
            Map.entry(Character.toString(119927), "{bold italic P}"),
            Map.entry(Character.toString(119928), "{bold italic Q}"),
            Map.entry(Character.toString(119929), "{bold italic R}"),
            Map.entry(Character.toString(119930), "{bold italic S}"),
            Map.entry(Character.toString(119931), "{bold italic T}"),
            Map.entry(Character.toString(119932), "{bold italic U}"),
            Map.entry(Character.toString(119933), "{bold italic V}"),
            Map.entry(Character.toString(119934), "{bold italic W}"),
            Map.entry(Character.toString(119935), "{bold italic X}"),
            Map.entry(Character.toString(119936), "{bold italic Y}"),
            Map.entry(Character.toString(119937), "{bold italic Z}"),
            Map.entry(Character.toString(119938), "{bold italic a}"),
            Map.entry(Character.toString(119939), "{bold italic b}"),
            Map.entry(Character.toString(119940), "{bold italic c}"),
            Map.entry(Character.toString(119941), "{bold italic d}"),
            Map.entry(Character.toString(119942), "{bold italic e}"),
            Map.entry(Character.toString(119943), "{bold italic f}"),
            Map.entry(Character.toString(119944), "{bold italic g}"),
            Map.entry(Character.toString(119945), "{bold italic h}"),
            Map.entry(Character.toString(119946), "{bold italic i}"),
            Map.entry(Character.toString(119947), "{bold italic j}"),
            Map.entry(Character.toString(119948), "{bold italic k}"),
            Map.entry(Character.toString(119949), "{bold italic l}"),
            Map.entry(Character.toString(119950), "{bold italic m}"),
            Map.entry(Character.toString(119951), "{bold italic n}"),
            Map.entry(Character.toString(119952), "{bold italic o}"),
            Map.entry(Character.toString(119953), "{bold italic p}"),
            Map.entry(Character.toString(119954), "{bold italic q}"),
            Map.entry(Character.toString(119955), "{bold italic r}"),
            Map.entry(Character.toString(119956), "{bold italic s}"),
            Map.entry(Character.toString(119957), "{bold italic t}"),
            Map.entry(Character.toString(119958), "{bold italic u}"),
            Map.entry(Character.toString(119959), "{bold italic v}"),
            Map.entry(Character.toString(119960), "{bold italic w}"),
            Map.entry(Character.toString(119961), "{bold italic x}"),
            Map.entry(Character.toString(119962), "{bold italic y}"),
            Map.entry(Character.toString(119963), "{bold italic z}")
        );

}
