package ru.danilakondr.gostproc;

import com.sun.star.beans.PropertyValue;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.lang.Exception;
import java.lang.RuntimeException;
import java.nio.file.Path;

/**
 * Главный класс постобработчика документов с использованием LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.1.0
 */
public class Application {
	private String docPath;
	private String docURL;
    private XDesktop xDesktop;
	private XTextDocument xDoc;
	private final XComponentContext xContext;
	private final XMultiComponentFactory xMCF;
	private boolean success = false;
	
	public Application(XComponentContext xContext) {
		this.xContext = xContext;
		this.xMCF = this.xContext.getServiceManager();
	};

	public static void main(String[] args) {
		XComponentContext xContext = null;

		try {
			xContext = LibreOffice.bootstrap();
		}
		catch (Exception e) {
			e.printStackTrace(System.err);
			System.exit(-1);
		}

		Application app = new Application(xContext);
		app.parseCommandLine(args);

		try {
			app.run();
			app.terminate();
		}
		catch (Exception e) {
			System.err.println(e.getMessage());
		}
	}

	/**
	 * Обработка аргументов командной строки.
	 *
	 * @param args массив с аргументами
	 */
	public void parseCommandLine(String[] args) {
		if (args.length < 1) {
			docPath = null;
		}
		else {
			docPath = args[0];
		}
	}

	/**
	 * Запуск приложения.
	 */
	public void run() throws Exception {
		this.createDesktop();
		this.loadDocument();

		new ParagraphStyleProcessor(xDoc).process();
		new PageStyleProcessor(xDoc).process();
		new MathFormulaProcessor(xDoc).process();
		new TableOfContentsProcessor(xDoc).process();
		new FirstPageStyleSetter(xDoc).process();

		this.success = true;
	}

	/**
	 * Попытка открытия файла по заданному в командной строке адресу.
	 * Устанавливает URL, по которому хранится обрабатываемый файл.
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
		XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xDoc);
		xStorable.store();
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