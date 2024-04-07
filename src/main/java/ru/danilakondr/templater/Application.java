package ru.danilakondr.templater;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.lang.Exception;
import java.lang.IllegalArgumentException;
import java.lang.RuntimeException;
import java.util.HashMap;
import java.util.List;

import com.sun.star.util.XCloseable;
import org.kohsuke.args4j.*;
import org.kohsuke.args4j.spi.MapOptionHandler;
import ru.danilakondr.templater.macros.*;
import ru.danilakondr.templater.processing.*;

/**
 * Главный класс постобработчика документов с использованием LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class Application {
	@Option(name="-t", aliases={"--template"}, usage="Template file", required = true)
	private String templatePath;

	@Option(name="-m", aliases={"--main-text"}, usage="Main text file", required = true)
	private String mainTextPath;

	@Option(name="-o", aliases={"--output"}, usage="Output file", required = true)
	private String outputPath;

	@Option(name="-M", aliases={"--macros"}, usage="Macros file")
	private String macroFile;

	@Option(name="-e", aliases={"--embed-fonts"}, usage="Embed fonts")
	private boolean embedFonts;

	@Option(name="-D", usage="Specify macro", handler=MapOptionHandler.class)
	private HashMap<String, String> macroOverrides;

	@Option(name="-h", aliases=	{"--help", "-?"})
	private boolean help;

	private final StringMacros stringMacros;
    private XDesktop xDesktop;
	private XTextDocument xDoc;
	private XComponentContext xContext;
	private XMultiComponentFactory xMCF;
	private boolean success = false;

	public Application() {
		this.xContext = null;
		this.xMCF = null;
		this.macroFile = null;
		this.stringMacros = new StringMacros();
	}

	public void setContext(XComponentContext xContext) {
		this.xContext = xContext;
		this.xMCF = this.xContext.getServiceManager();
	}

	public static void main(String[] args) {
		Application app = new Application();
		CmdLineParser parser = new CmdLineParser(app);
		try {
			parser.parseArgument(args);
		}
		catch (CmdLineException e) {
			System.err.println(e.getMessage());
			System.exit(-1);
		}

		if (app.help) {
			System.out.println("Usage:");
			System.out.print("templater ");
			parser.printSingleLineUsage(System.out);
			System.out.println();
			System.out.println();
			parser.printUsage(System.out);
			System.exit(0);
		}

		XComponentContext xContext = null;
		try {
			xContext = LibreOffice.bootstrap();
			app.setContext(xContext);
		}
		catch (Exception e) {
			e.printStackTrace(System.err);
			System.exit(-1);
		}

		int status = 0;
		try {
			app.run();
		}
		catch (Exception e) {
			e.printStackTrace(System.err);
			status = 1;
		}
		finally {
			app.terminate();
			System.exit(status);
		}
	}

	/**
	 * Запуск приложения.
	 */
	public void run() throws Exception {
		this.createDesktop();

		String mainTextURL = getURI(mainTextPath, true);
		if (mainTextURL == null) {
			throw new FileNotFoundException(mainTextPath);
		}

		if (macroFile != null) {
			if (new File(macroFile).exists()) {
				System.err.println("Loading macros from file...");
				stringMacros.loadFromFile(macroFile);
			}
			else {
				System.err.printf("File %s not found, skipping\n", macroFile);
			}
		}

		stringMacros.loadFromMap(macroOverrides);

		fixFormulasAlignmentInMainFile(mainTextURL);

		this.loadTemplate();

		MacroSubstitutor substitutor = new MacroSubstitutor(xDoc);
		substitutor.substitute(new MainTextIncludeSubstitutor(), mainTextURL);
		for (int i = 0; i < 16; i++)
			substitutor.substitute(new DocumentIncludeSubstitutor(), null);
		substitutor
				.substitute(new StringMacroSubstitutor(), stringMacros)
				.substitute(new TableOfContentsInserter(), null);

		TextDocument proc = new TextDocument(xDoc);
		proc
				.processFormulas(new MathFormulaFixProcessor(), new ProgressCounter("Fixing formulas"))
				.processFormulas(new ObjectAligner(), new ProgressCounter("Aligning formulas properly"))
				.processParagraphs(new NumberingStyleProcessor(), new ProgressCounter("Processing numbering style of paragraphs"))
				.processImages(new ImageWidthFixProcessor(), new ProgressCounter("Fixing image widths"))
				.processImages(new ObjectAligner(), new ProgressCounter("Fixing image alignments"));
		ProgressCounter tablesCnt = new ProgressCounter("Processing tables");
		List<XTextSection> tables = proc
				.streamSections()
				.filter(s -> UnoRuntime.queryInterface(XNamed.class, s).getName().startsWith("tbl:"))
				.toList();
		tables.forEach(x -> {
			try {
				String name  = UnoRuntime.queryInterface(XNamed.class, x).getName();
				tablesCnt.setString(String.format("Processing tables inside section %s", name));
				tablesCnt.setShowCurrent(false);
				proc.processTablesInsideRange(
						x.getAnchor(),
						new TableStyleSetter(),
						tablesCnt
				);
			} catch (com.sun.star.uno.Exception e) {
				throw new RuntimeException(e);
			}
		});

		proc.updateAllIndexes();

		this.success = true;
	}

	/**
	 * Исправляет выравнивание формул в главном файле.
	 *
	 * @param mainTextURL URL-адрес файла
	 * @since 0.3.2
	 */
	private void fixFormulasAlignmentInMainFile(String mainTextURL) throws Exception {
		XTextDocument xMainDoc = this.loadFile(mainTextURL);
		XCloseable xMainDocCloseable = UnoRuntime.queryInterface(XCloseable.class, xMainDoc);
		XStorable xMainDocStorable = UnoRuntime.queryInterface(XStorable.class, xMainDoc);

		TextDocument mainDoc = new TextDocument(xMainDoc);
		mainDoc.processFormulas((o, d) -> {
			XPropertySet xContentProps = UnoRuntime.queryInterface(XPropertySet.class, o);
			try {
				xContentProps.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
			}
			catch (Exception ignored) {}
		}, new ProgressCounter("Processing formula in main file"));

		xMainDocStorable.store();
		xMainDocCloseable.close(true);
	}

	/**
	 * Создаёт &laquo;рабочий стол&raquo; LibreOffice.
	 */
	private void createDesktop() throws java.lang.Exception {
		Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext);
		xDesktop = UnoRuntime.queryInterface(XDesktop.class, oDesktop);
	}

	private String getURI(String path, boolean read) {
		File f = new File(path).getAbsoluteFile();

		if (read && !f.exists())
			return null;

		return f.toPath().toUri().toString();
	}

	/**
	 * Загружает текстовый документ.
	 *
	 * @param url URL-адрес файла
	 * @return текстовый документ
	 * @since 0.3.2
	 */
	private XTextDocument loadFile(String url) throws Exception {
		XComponentLoader xCompLoader = UnoRuntime
				.queryInterface(XComponentLoader.class, xDesktop);

		PropertyValue[] props = new PropertyValue[1];
		props[0] = new PropertyValue();
		props[0].Name = "Hidden";
		props[0].Value = Boolean.TRUE;

		XComponent xComp = xCompLoader.loadComponentFromURL(url, "_blank", 0, props);
		XServiceInfo xServiceInfo = UnoRuntime.queryInterface(XServiceInfo.class, xComp);
		if (!xServiceInfo.supportsService("com.sun.star.text.TextDocument"))
			throw new IllegalArgumentException("Invalid format");
		XTextDocument doc = UnoRuntime.queryInterface(XTextDocument.class, xComp);

		return doc;
	}

	/**
	 * Загружает шаблон.
	 *
	 * @since 0.1.0
	 */
	private void loadTemplate() throws Exception {
		String templateURL = getURI(templatePath, true);
		if (templateURL == null)
			throw new FileNotFoundException(templatePath);

		this.xDoc = loadFile(templateURL);
	}

	/**
	 * Сохраняет полученный документ. Если нужно, внедряет шрифты.
	 *
	 * @since 0.1.0
	 */
	private void saveDocument() throws Exception {
		String outputURL = getURI(outputPath, false);
		XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xDoc);

		if (embedFonts) {
			XMultiServiceFactory xMSF = UnoRuntime
					.queryInterface(XMultiServiceFactory.class, xDoc);
			XPropertySet xDocSettings = UnoRuntime.queryInterface(
					XPropertySet.class,
					xMSF.createInstance("com.sun.star.text.DocumentSettings")
			);
			xDocSettings.setPropertyValue("EmbedFonts", true);
			xDocSettings.setPropertyValue("EmbedOnlyUsedFonts", true);
		}
		xStorable.storeAsURL(outputURL, new PropertyValue[0]);
	}

	/**
	 * Закрывает документ.
	 *
	 * @since 0.3.3
	 */
	private void closeDocument() throws Exception {
		XCloseable xCloseable = UnoRuntime
				.queryInterface(XCloseable.class, xDoc);
		xCloseable.close(true);
	}

	/**
	 * Завершает приложение.
	 *
	 * @since 0.1.0
	 */
	public void terminate() {
		if (this.success) {
			try {
				this.saveDocument();
			} catch (Exception e) {
				System.err.println(e.getMessage());
			}
		}

		try {
			this.closeDocument();
		}
		catch (Exception e) {
			throw new RuntimeException(e);
		}
		finally {
			System.out.println("Terminated");
		}
	}
}