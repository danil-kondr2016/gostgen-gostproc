package ru.danilakondr.gostproc;

import com.sun.star.beans.PropertyValue;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;

import java.io.File;
import java.lang.Exception;
import java.lang.RuntimeException;
import java.nio.file.Path;

import org.kohsuke.args4j.*;
import ru.danilakondr.gostproc.fonts.FontTriple;
import ru.danilakondr.gostproc.fonts.FontTripleHandler;
import ru.danilakondr.gostproc.processing.*;

/**
 * Главный класс постобработчика документов с использованием LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class Application {
	@Argument(usage="Input file", metaVar="INPUT")
	private String docPath;
	@Option(name="-o", aliases={"--out"}, usage="Output file", metaVar="OUTPUT")
	private String outPath;

	@Option(name="-h", aliases={"--help", "-?"})
	private boolean help;

	private String docURL;
    private XDesktop xDesktop;
	private XTextDocument xDoc;
	private XComponentContext xContext;
	private XMultiComponentFactory xMCF;
	private boolean success = false;

	public Application() {
		this.xContext = null;
		this.xMCF = null;
	}

	private void setContext(XComponentContext xContext) {
		this.xContext = xContext;
		this.xMCF = this.xContext.getServiceManager();
	}

	public static void main(String[] args) {
		OptionHandlerRegistry
				.getRegistry()
				.registerHandler(FontTriple.class, FontTripleHandler.class);

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
		this.loadDocument();

		new DocumentIncluder(xDoc).process();
		new MathFormulaProcessor(xDoc).process();
		new TableOfContentsProcessor(xDoc).process();

		this.success = true;
	}

	/**
	 * Попытка открытия файла по заданному в командной строке адресу.
	 * Определяет URL, по которому хранится обрабатываемый файл.
	 */
	private void tryToOpen() throws Exception {
		if (docPath != null) {
			File f = new File(docPath);
			if (!f.exists()) {
				throw new RuntimeException("Cannot open file " + f);
			}

			Path p = f.toPath();
			this.docURL = p.toUri().toString();
		}

		if (docURL == null)
			throw new Exception("File has not been specified");
	}

	/**
	 * Создаёт &laquo;рабочий стол&raquo; LibreOffice.
	 */
	private void createDesktop() throws java.lang.Exception {
		Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext);
		xDesktop = UnoRuntime.queryInterface(XDesktop.class, oDesktop);
	}

	/**
	 * Открывает документ.
	 */
	private void loadDocument() throws Exception {
		tryToOpen();

		XComponentLoader xCompLoader = UnoRuntime
				.queryInterface(XComponentLoader.class, xDesktop);

		PropertyValue[] props = new PropertyValue[1];
		props[0] = new PropertyValue();
		props[0].Name = "Hidden";
		props[0].Value = Boolean.TRUE;

		XComponent xComp = xCompLoader.loadComponentFromURL(docURL, "_blank", 0, props);
		xDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);
	}

	/**
	 * Закрывает документ.
	 */
	private void closeDocument() throws Exception {
        String outURL;
        if (outPath != null) {
			outURL = Path.of(outPath)
					.toAbsolutePath()
					.toUri()
					.toString();
		} else {
			outURL = null;
		}

		XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xDoc);
		if (outURL == null) {
			xStorable.store();
		} else {
			PropertyValue[] props = new PropertyValue[0];
			xStorable.storeAsURL(outURL, props);
		}
	}

	/**
	 * Завершает приложение.
	 */
	public void terminate() {
		if (this.success) {
			try {
				this.closeDocument();
			} catch (Exception e) {
				System.err.println(e.getMessage());
			}
		}

		xDesktop.terminate();
	}
}