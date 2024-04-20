package ru.danilakondr.templater.app;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.text.*;
import com.sun.star.frame.*;
import com.sun.star.uno.*;
import com.sun.star.lang.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.Exception;
import java.lang.IllegalArgumentException;
import java.lang.RuntimeException;
import java.net.URI;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.util.Map;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.sun.star.util.XCloseable;
import ru.danilakondr.templater.macros.*;
import ru.danilakondr.templater.processing.*;
import ru.danilakondr.templater.progress.DefaultProgressInformer;

/**
 * Главный класс постобработчика документов с использованием LibreOffice.
 *
 * @author Данила А. Кондратенко
 * @since 0.5.0, 0.1.0
 */
public class Templater {
	private String templatePath;
	private String mainTextPath;
	private String outputPath;

	private transient String templateURL;
	private transient String mainTextURL;
	private transient String outputURL;

	private boolean shouldEmbedFonts;
	private boolean shouldOverwrite;
	private boolean verbose;

	private final StringMacros stringMacros;

    private XDesktop xDesktop;
	private XTextDocument xDoc;
	private XComponentContext xContext;
	private XMultiComponentFactory xMCF;

	private final DefaultProgressInformer informer;

	public Templater() {
		this.xContext = null;
		this.xMCF = null;
		this.stringMacros = new StringMacros();
		this.informer = new DefaultProgressInformer("Doing");
	}

	public void setContext(XComponentContext xContext) {
		this.xContext = xContext;
		this.xMCF = this.xContext.getServiceManager();
	}

	public void setOutputPath(String outputPath) {
		this.outputPath = outputPath;
		this.outputURL = getURI(outputPath);
	}

	public void setTemplatePath(String templatePath) {
		this.templatePath = templatePath;
		this.templateURL = getURI(templatePath);
	}

	public void setMainTextPath(String mainTextPath) {
		this.mainTextPath = mainTextPath;
		this.mainTextURL = getURI(mainTextPath);
	}

	public void setShouldOverwrite(boolean shouldOverwrite) {
		this.shouldOverwrite = shouldOverwrite;
	}

	public void setShouldEmbedFonts(boolean shouldEmbedFonts) {
		this.shouldEmbedFonts = shouldEmbedFonts;
	}

	public void loadMacrosFromFile(String path) throws IOException {
		stringMacros.loadFromFile(path);
	}

	public void loadMacrosFromMap(Map<String, String> macros) {
		stringMacros.loadFromMap(macros);
	}

	public void setVerbose(boolean verbose) {
		this.verbose = verbose;
		this.informer.setSilent(!verbose);
	}

	public void processDocument() throws Exception {
		this.checkFiles();
		this.createDesktop();

		fixObjectAlignmentInFile(mainTextURL);
		this.loadTemplate();

		this.substituteMacros();
		this.fixDocument();
		this.applyCounters();
	}

	private void substituteMacros() throws Exception {
		informer.setProgressString("Substituting macros");
		informer.inform(-1, -1);

		MacroSubstitutor substitutor = new MacroSubstitutor(xDoc);
		substitutor.substitute(new MainTextIncludeSubstitutor(mainTextURL));
		substitutor.substitute(new DocumentIncludeSubstitutor());
		substitutor.substitute(new StringMacroSubstitutor(stringMacros));
		substitutor.substitute(new TableOfContentsInserter());
	}

	private void fixDocument() throws Exception {
		TextDocument document = new TextDocument(xDoc);

		informer.setProgressString("Fixing formulas");
		document.processFormulas(new MathFormulaFixProcessor(), informer);

		informer.setProgressString("Aligning formulas properly");
		document.processFormulas(new SingleObjectAligner(), informer);

		informer.setProgressString("Processing numbering style of paragraphs");
		document.processParagraphs(new NumberingStyleProcessor(), informer);

		informer.setProgressString("Fixing image widths");
		document.processImages(new ImageSizeFixProcessor(), informer);

		informer.setProgressString("Fixing image alignments");
		document.processImages(new SingleObjectAligner(), informer);

		informer.setProgressString("Setting table styles");
		document.processTables(new TableStyleSetter(), informer);

		informer.setProgressString("Updating styles");
		document.updateAllIndexes(informer);
	}

	private void applyCounters() throws Exception {
		informer.setProgressString("Applying counters");
		informer.inform(-1, -1);

		new MacroSubstitutor(xDoc)
				.substitute(new StringMacroSubstitutor(
						DocumentCounter.getCounter(xDoc)));
	}

	private boolean isFileNotExists(String url) {
		try {
			File f = Path.of(new URI(url)).toFile();
			return !f.exists();
		}
		catch (URISyntaxException e) {
			return true;
		}
	}

	private void checkFiles() throws Exception {
		if (templatePath == null)
			throw new IllegalArgumentException("Template file has not been specified");
		if (mainTextPath == null)
			throw new IllegalArgumentException("Main text file has not been specified");
		if (outputPath == null)
			throw new IllegalArgumentException("Output file has not been specified");

		if (isFileNotExists(templateURL))
			throw new FileNotFoundException(templatePath);
		if (isFileNotExists(mainTextURL))
			throw new FileNotFoundException(mainTextPath);

		if (mainTextURL.equals(outputURL))
			throw new IllegalArgumentException("Main text file and output file cannot be the same file");
		if (templateURL.equals(outputURL))
			throw new IllegalArgumentException("Template and output file cannot be the same file");
		if (templateURL.equals(mainTextURL))
			throw new IllegalArgumentException("Template and main text file cannot be the same file");

        if (Path.of(new URI(outputURL)).toFile().exists() && !shouldOverwrite) {
			askForOverwrite(outputPath);
		}
	}

	private void askForOverwrite(String outputPath) {
		boolean selected = false;
		try (Scanner sc = new Scanner(System.in)) {
			System.out.printf("Overwrite %s (Y/N)? ", outputPath);
			while (!selected) {
				String choice = sc.nextLine().trim();

				if (choice.equalsIgnoreCase("y")) {
					selected = true;
					shouldOverwrite = true;
				} else if (choice.equalsIgnoreCase("n")) {
					selected = true;
					shouldOverwrite = false;
				} else {
					System.out.println("Invalid choice: " + choice);
				}
			}
			System.out.println();
		}
		catch (Exception e) {
			e.printStackTrace(System.err);
		}

		if (!shouldOverwrite) {
			throw new RuntimeException("Could not overwrite existing file");
		}
	}

	/**
	 * Исправляет выравнивание формул и изображений в главном файле.
	 *
	 * @param mainTextURL URL-адрес файла
	 * @since 0.3.2
	 */
	private void fixObjectAlignmentInFile(String mainTextURL) throws Exception {
		XTextDocument xMainDoc = this.loadFile(mainTextURL);
		XCloseable xMainDocCloseable = UnoRuntime.queryInterface(XCloseable.class, xMainDoc);
		XStorable xMainDocStorable = UnoRuntime.queryInterface(XStorable.class, xMainDoc);

		informer.setProgressString("Processing formulas in main file");
		TextDocument mainDoc = new TextDocument(xMainDoc);
		mainDoc.processFormulas((o, d) -> {
			XPropertySet xContentProps = UnoRuntime.queryInterface(XPropertySet.class, o);
			try {
				xContentProps.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
			}
			catch (Exception ignored) {}
		}, informer);

		informer.setProgressString("Processing images in main file");
		mainDoc.processImages((o, d) -> {
			XPropertySet xContentProps = UnoRuntime.queryInterface(XPropertySet.class, o);
			try {
				xContentProps.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
			}
			catch (Exception ignored) {}
		}, informer);

		xMainDocStorable.store();
		xMainDocCloseable.close(true);
	}

	/**
	 * Создаёт &laquo;рабочий стол&raquo; LibreOffice.
	 *
	 * @since 0.1.0
	 */
	private void createDesktop() throws java.lang.Exception {
		Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext);
		xDesktop = UnoRuntime.queryInterface(XDesktop.class, oDesktop);
	}

	private String getURI(String path) {
		if (path == null)
			return null;

		File f = new File(path).getAbsoluteFile();
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
		if (xServiceInfo == null)
			throw new IllegalArgumentException("Failed to create XComponent");

		if (!xServiceInfo.supportsService("com.sun.star.text.TextDocument")) {
			XCloseable xCloseable = UnoRuntime.queryInterface(
					XCloseable.class, xComp
			);
			if (xCloseable == null)
				xComp.dispose();
			else
				xCloseable.close(false);

			throw new IllegalArgumentException("Invalid format of document");
		}

        return UnoRuntime.queryInterface(XTextDocument.class, xComp);
	}

	/**
	 * Загружает шаблон.
	 *
	 * @since 0.1.0
	 */
	private void loadTemplate() throws Exception {
		this.xDoc = loadFile(templateURL);
	}

	/**
	 * Сохраняет полученный документ. Если нужно, внедряет шрифты.
	 *
	 * @since 0.1.0
	 */
	public void saveDocument() throws Exception {
		informer.setProgressString("Saving document");
		informer.inform(-1, -1);

		XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xDoc);

		if (shouldEmbedFonts) {
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
	 * Экспортирует готовый файл в PDF при условии, если имеется соответствующий
	 * аргумент.
	 * <p>
	 * В названии файла расширение заменяется на PDF.
	 *
	 * @since 0.4.2
	 */
	public void generatePDF() throws Exception {
		informer.setProgressString("Generating PDF file");
		informer.inform(-1, -1);

		String pdfURL = outputURL;

		Pattern ext = Pattern.compile("(.*)\\.(.*?)$");
		Matcher m = ext.matcher(pdfURL);

		if (m.matches())
			pdfURL = m.replaceAll("$1.pdf");
		else
			pdfURL += ".pdf";

		XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xDoc);

		PropertyValue[] propertyValues = new PropertyValue[2];

		propertyValues[0] = new PropertyValue();
		propertyValues[0].Name = "Overwrite";
		propertyValues[0].Value = Boolean.TRUE;

		propertyValues[1] = new PropertyValue();
		propertyValues[1].Name = "FilterName";
		propertyValues[1].Value = "writer_pdf_Export";

		xStorable.storeToURL(pdfURL, propertyValues);
	}

	/**
	 * Закрывает документ.
	 *
	 * @since 0.3.3
	 */
	public void closeDocument() throws Exception {
		if (xDoc == null)
			return;

		XCloseable xCloseable = UnoRuntime
				.queryInterface(XCloseable.class, xDoc);
		xCloseable.close(true);
	}
}
