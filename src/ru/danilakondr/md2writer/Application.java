package ru.danilakondr.md2writer;

import com.sun.star.beans.PropertyValue;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.comp.helper.Bootstrap;

import java.io.File;
import java.lang.Exception;
import java.lang.RuntimeException;
import java.nio.file.Path;

public class Application {
	private final String docPath;
	private String docURL;
    private XDesktop xDesktop;
	private XTextDocument xDoc;
	
	public Application(String docPath) {
		this.docPath = docPath;
	}

	public static void main(String[] args) {
		Application app = new Application(args[0]);

		try {
			app.run();
		}
		catch (Exception e) {
			System.err.println(e.getMessage());
		}
	}

	public void run() throws Exception {
		this.bootstrap();
		this.loadDocument();

		new TableOfContentsProcessor(xDoc).process();
		new MathFormulaProcessor(xDoc).process();
		this.closeDocument();
	}

	private void tryToOpen() throws RuntimeException {
		File f = new File(docPath);
		if (!f.exists()) {
			throw new RuntimeException("Cannot open file " + f);
		}

		Path p = f.toPath();
		this.docURL = p.toUri().toString();
	}
	
	private void bootstrap() throws java.lang.Exception {
        XComponentContext xContext = Bootstrap.bootstrap();
        XMultiComponentFactory xMCF = xContext.getServiceManager();
		Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext);
		xDesktop = UnoRuntime.queryInterface(XDesktop.class, oDesktop);
	}
	
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

	private void closeDocument() throws Exception {
		XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xDoc);
		xStorable.store();

		xDesktop.terminate();
	}
}