package ru.danilakondr.gostproc;

import com.sun.star.beans.PropertyValue;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.ui.dialogs.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.comp.helper.Bootstrap;

import javax.swing.text.Style;
import java.io.File;
import java.lang.Exception;
import java.lang.RuntimeException;
import java.nio.file.Path;

public class Application {
	private String docPath;
	private String docURL;
    private XDesktop xDesktop;
	private XTextDocument xDoc;
	private XComponentContext xContext;
	private XMultiComponentFactory xMCF;
	private boolean success = false;
	
	public Application() {};

	public static void main(String[] args) {
		Application app = new Application();
		app.parseCommandLine(args);

		try {
			app.run();
		}
		catch (Exception e) {
			System.err.println(e.getMessage());
		}
		finally {
			app.terminate();
		}
	}

	public void parseCommandLine(String[] args) {
		if (args.length < 1) {
			docPath = null;
		}
		else {
			docPath = args[0];
		}
	}

	public void run() throws Exception {
		this.bootstrap();
		this.loadDocument();

		new StyleProcessor(xDoc).process();
		new TableOfContentsProcessor(xDoc).process();
		new MathFormulaProcessor(xDoc).process();

		this.success = true;
		this.closeDocument();
	}

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
			throw new RuntimeException("File has not been specified");
	}

	private void bootstrap() throws java.lang.Exception {
        xContext = Bootstrap.bootstrap();
		xMCF = xContext.getServiceManager();
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
	}

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