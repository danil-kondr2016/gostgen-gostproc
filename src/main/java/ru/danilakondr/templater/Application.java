package ru.danilakondr.templater;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.lang.Exception;

import org.kohsuke.args4j.*;
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

	@Option(name="-h", aliases=	{"--help", "-?"})
	private boolean help;

	private final Macros macros;
    private XDesktop xDesktop;
	private XTextDocument xDoc;
	private XComponentContext xContext;
	private XMultiComponentFactory xMCF;
	private boolean success = false;

	public Application() {
		this.xContext = null;
		this.xMCF = null;
		this.macroFile = null;
		this.macros = new Macros();
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
			System.out.print("postproc ");
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

		this.loadTemplate();

		new DocumentIncluder(xDoc, mainTextURL).process();
		new MathFormulaProcessor(xDoc).process();
		if (macroFile != null) {
			if (new File(macroFile).exists()) {
				macros.loadFromFile(macroFile);
			}
			else {
				System.err.printf("File %s not found, skipping\n", macroFile);
			}
		}
		new MacroProcessor(xDoc, macros).process();
		new ImageWidthFixer(xDoc).process();
		new NumberingStyleProcessor(xDoc).process();
		new TableStyleProcessor(xDoc).process();
		new TableOfContentsInserter(xDoc).process();

		this.success = true;
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

	private void loadTemplate() throws Exception {
		String templateURL = getURI(templatePath, true);
		if (templateURL == null)
			throw new FileNotFoundException(templatePath);

		XComponentLoader xCompLoader = UnoRuntime
				.queryInterface(XComponentLoader.class, xDesktop);

		PropertyValue[] props = new PropertyValue[1];
		props[0] = new PropertyValue();
		props[0].Name = "Hidden";
		props[0].Value = Boolean.TRUE;

		XComponent xComp = xCompLoader.loadComponentFromURL(templateURL, "_blank", 0, props);
		xDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);
	}

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
	 * Завершает приложение.
	 *
	 * FIXME: приложение закрывается вместе со всем Либреофисом.
	 */
	public void terminate() {
		if (this.success) {
			try {
				this.saveDocument();
			} catch (Exception e) {
				System.err.println(e.getMessage());
			}
		}

		xDesktop.terminate();
	}
}