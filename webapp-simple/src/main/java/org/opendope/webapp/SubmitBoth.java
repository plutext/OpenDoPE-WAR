/*
 *  Copyright 2010-13, Plutext Pty Ltd.
 *   
 *  This file is part of OpenDoPE Java simple webapp.

    OpenDoPE Java simple webapp is licensed under the Apache License, 
    Version 2.0 (the "License"); 
    you may not use this file except in compliance with the License. 

    You may obtain a copy of the License at 

        http://www.apache.org/licenses/LICENSE-2.0 

    Unless required by applicable law or agreed to in writing, software 
    distributed under the License is distributed on an "AS IS" BASIS, 
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
    See the License for the specific language governing permissions and 
    limitations under the License.

 */
package org.opendope.webapp;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

import javax.annotation.PostConstruct;
import javax.servlet.ServletConfig;
import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.WebApplicationException;
import javax.ws.rs.core.Context;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;
import javax.ws.rs.core.Response.Status;
import javax.ws.rs.core.StreamingOutput;
import javax.xml.bind.JAXBContext;

import org.apache.commons.io.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.convert.out.html.AbstractHtmlExporter;
import org.docx4j.convert.out.html.AbstractHtmlExporter.HtmlSettings;
import org.docx4j.convert.out.html.HtmlExporterNG2;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.model.datastorage.BindingHandler;
import org.docx4j.model.datastorage.CustomXmlDataStoragePartSelector;
import org.docx4j.model.datastorage.OpenDoPEHandler;
import org.docx4j.model.datastorage.OpenDoPEIntegrity;
import org.docx4j.model.datastorage.RemovalHandler;
import org.docx4j.model.datastorage.RemovalHandler.Quantifier;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.LoadFromZipNG;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.docx4j.wml.Document;
import org.glassfish.jersey.media.multipart.FormDataContentDisposition;
import org.glassfish.jersey.media.multipart.FormDataParam;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.docx4j.services.client.Converter;
import org.docx4j.services.client.ConverterHttp;
import org.docx4j.services.client.Format;
import org.docx4j.toc.TocException;


@Path("/both")  // must match form action
public class SubmitBoth {

	final static String KEY_XML = "xmlfile";
	final static String KEY_DOCX = "docx";

	public static JAXBContext context = org.docx4j.jaxb.Context.jc; 
	
	private static Logger log = LoggerFactory.getLogger(SubmitBoth.class);		

    @Context ServletConfig servletConfig;
    //@Context ServletContext servletContext;
    
    static String hyperlinkStyleId;
    static String htmlImageTargetUri;
    static String htmlImageDirPath;
    
    static String documentServicesEndpoint;
    
    @PostConstruct
    public void readInitParams() {
    	
    	// This works with and without class extends Application
       
    	log.info( servletConfig.getInitParameter("HyperlinkStyleId") );
    	log.info( servletConfig.getInitParameter("HtmlImageTargetUri") );
    	log.info( servletConfig.getInitParameter("HtmlImageDirPath") );
    	
        hyperlinkStyleId = servletConfig.getInitParameter("HyperlinkStyleId");
        htmlImageTargetUri = servletConfig.getInitParameter("HtmlImageTargetUri");
        htmlImageDirPath = servletConfig.getInitParameter("HtmlImageDirPath");

    	log.info( servletConfig.getInitParameter("DocumentServicesEndpoint") );
        documentServicesEndpoint = servletConfig.getInitParameter("DocumentServicesEndpoint");
        
    }     
	
	
    private static Set<String> cssWhiteList = null;
	private void initWhitelist() {
		
		if (cssWhiteList!=null) return; // done already
		
		try {
			cssWhiteList = new HashSet<String>();
			List lines = IOUtils.readLines(servletConfig.getServletContext().getResourceAsStream("/WEB-INF/CSS-WhiteList.txt"), "UTF-8" );
			for (Object o : lines) {
				String line = ((String)o).trim();
				if (line.length()>0 && !line.startsWith("#")) {
					cssWhiteList.add(line);
				}
			}
			XHTMLImporterImpl.setCssWhiteList(cssWhiteList);
		} catch (IOException e) {
			log.warn("CSS-WhiteList not found", e);
		}
		
		
        // demo font mapping
//		RFonts rfonts = org.docx4j.jaxb.Context.getWmlObjectFactory().createRFonts();
//		rfonts.setAscii("Comic Sans MS");
//        XHTMLImporterImpl.addFontMapping("Century Gothic", rfonts);
		
		
	}
    
    
	/**
	 * Display a form in which user can provide OpenDoPE docx template,
	 * and xml data file, for processing.  Useful for testing.
	 * The URL will be something like http://localhost:8080/OpenDoPE-simple/service/both
	 */
	@GET
	@Produces("text/html")
	public String uploadForm() {
		return "<html><head>" 
//				+"		<meta http-equiv=\"Content-type\" content=\"text/html;charset=UTF-8\" /> <!--  IE and Firefox need this -->"	
				+"<title>OpenDoPE injection</title></head>"
				+ "<body>"
				
//				+ "<script type=\"text/javascript\">"
//				+ "function OnSubmitForm()"
//				+ "{"
//				+ "  if(document.uploadForm.format.checked == true) {"
//				+ "    document.uploadForm.action =\"both?format=html\";"
//				+ "  } else  {"
//				+ "    document.uploadForm.action =\"both\";"
//				+ "  }"
//				+ "  return true;"
//				+ "}"
//				+ "</script>"	
				
				+ "<form name=\"uploadForm\" action=\"both\" " 
				+		" onsubmit=\"return OnSubmitForm();\" "
				+ 		" method=\"post\" enctype=\"multipart/form-data\">"
				+ "<label for=\"" + KEY_XML + "\">XML data file:</label> "
				+ "<input type=\"file\" name=\"" + KEY_XML + "\" id=\"" + KEY_XML + "\" /> "
				+ "<br/>"
				+ "<label for=\"" + KEY_DOCX + "\">docx template:</label> "
				+ "<input type=\"file\" name=\"" + KEY_DOCX + "\" id=\"" + KEY_DOCX + "\" /> "
				+ "<br/>"
				+ "<input type=\"radio\" name=\"format\" value=\"docx\" >as DOCX</input>"
				+ "<input type=\"radio\" name=\"format\" value=\"html\" >as HTML</input>"
				+ "<input type=\"radio\" name=\"format\" value=\"pdf\" >as PDF</input>"
				+ "<br/>"
				+ "<input type=\"checkbox\" name=\"processTocField\"  value=\"true\">process ToC field</input>"
				+ "<input type=\"checkbox\" name=\"TocPageNumbers\"  value=\"true\">include page numbers</input>"
				+ "<br/>"
				+ "<input type=\"submit\" name=\"submit\" value=\"Submit\" /> "
				+ "</form>" + "</body></html>";
	}
	
	/*
	 * From http://wiki.apache.org/tomcat/FAQ/CharacterEncoding
	 * 
	 * Default Encoding for POST

		ISO-8859-1 is defined as the default character set for HTTP request and response 
		bodies in the servlet specification (request encoding: section 4.9 for 
		spec version 2.4, section 3.9 for spec version 2.5; response encoding: 
		section 5.4 for both spec versions 2.4 and 2.5). This default is historical: 
		it comes from sections 3.4.1 and 3.7.1 of the HTTP/1.1 specification.
		
		Some notes about the character encoding of a POST request:
		
		Section 3.4.1 of HTTP/1.1 states that recipients of an HTTP message must respect 
		the character encoding specified by the sender in the Content-Type header if the
		 encoding is supported. A missing character allows the recipient to "guess" 
		 what encoding is appropriate.
		 
		Most web browsers today do not specify the character set of a request, even when 
		it is something other than ISO-8859-1. This seems to be in violation of the HTTP 
		specification. Most web browsers appear to send a request body using the encoding 
		of the page used to generate the POST (for instance, the <form> element came from 
		a page with a specific encoding... it is that encoding which is used to submit the 
		POST data for that form).
		
		-----------------------------------
		
		If you suspect character encoding problems, first check that UTF-8 is being used
		in the build.
		
			Eclipse: 
			
			Window > Preferences > General > Workspace
			Check that Text file encoding is UTF-8.
			
			go into Project properties > Resource.
			Check that "text file encoding" is UTF-8 (either Inherited from container, or other).
			
		In Maven, you need <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
		(on docx4j etc as well).
		
		If there are still problems, you might consider an XML data file containing UTF-8 characters ϣ • € ₩
 
			 -Dfile.encoding=UTF-8 in Tomcat's setenv.bat doesn't seem to matter
			
			&lt;meta http-equiv="Content-type" content="text/html;charset=UTF-8" /&gt;   not required in escaped XHTML
			
			<meta http-equiv =\"Content-type\" content=\"text/html;charset =UTF-8\" />  not required in uploadForm() form above
			
			<?xml version='1.0' encoding='UTF-8'?> not even required on input xml file
			
			No need for org.apache.catalina.filters.SetCharacterEncodingFilter

		
	 */

	@POST
	@Consumes(MediaType.MULTIPART_FORM_DATA)
	@Produces( {"application/vnd.openxmlformats-officedocument.wordprocessingml.document" , 
				"text/html"})
	public Response processForm(
//			@Context ServletContext servletContext, 
//			@Context  HttpServletRequest servletRequest,
			@FormDataParam("xmlfile") InputStream xis,
			@FormDataParam("xmlfile") FormDataContentDisposition xmlFileDetail,
			@FormDataParam("docx") InputStream docxInputStream,
			@FormDataParam("docx") FormDataContentDisposition docxFileDetail,
			@FormDataParam("format") String format,
			@FormDataParam("processTocField") boolean processTocField,
			@FormDataParam("TocPageNumbers") boolean tocPageNumbers
			) throws Docx4JException, IOException {
		
//		log.info("servletRequest.getCharacterEncoding(): " + servletRequest.getCharacterEncoding());
//		
//		log.warn("platformt: " + getDefaultEncoding() );
//		
//		log.info("requested format: " + format);
//		
//		byte[] bytes = IOUtils.toByteArray(xis);
//		File f = new File( System.getProperty("java.io.tmpdir") + "/xml.xml");
//		FileUtils.writeByteArrayToFile(f, bytes); 
//		log.info("Saved: " + f.getAbsolutePath());

//		log.info("gip" + servletContext.getInitParameter("HtmlImageTargetUri") );
//		java.util.Enumeration penum = servletContext.getInitParameterNames();
//		for ( ; penum.hasMoreElements() ;) {
//			String name = (String)penum.nextElement();
//			log.info( name + "=" +  servletContext.getInitParameter(name) );
//		}
		
		// For XHTML Import, only honour CSS on the white list
		initWhitelist();
		
		// XHTML Import: no need to add a mapping, if docx defaults are to be used,
		// for fonts in docx4j's MicrosoftFonts.xml.
		// Helvetica is not there.
		XHTMLImporterImpl.addFontMapping("helvetica", "Helvetica");		
		
		final WordprocessingMLPackage wordMLPackage;

		String dataname = getFileName(xmlFileDetail.getFileName());
		
		WordprocessingMLPackage tmpPkg=null;
			
		String docxname = getFileName(docxFileDetail.getFileName());
		
		LoadFromZipNG loader = new LoadFromZipNG();
		try {
			tmpPkg = (WordprocessingMLPackage)loader.get(docxInputStream );
		} catch (Exception e) {
			throw new WebApplicationException(
					new Docx4JException("Error reading docx file (is this a valid docx?)"), 
					Status.BAD_REQUEST);
		}
			
		
		
		// Now we should have both our docx and xml
		if (tmpPkg==null) {
			throw new WebApplicationException(
					new Docx4JException("No docx file provided"), 
					Status.BAD_REQUEST);
		}
		if (xis==null) {
			throw new WebApplicationException(
					new Docx4JException("No xml data file provided"), 
					Status.BAD_REQUEST);
		}
		wordMLPackage = tmpPkg;
		String filePrefix = docxname + "_" + dataname + "_" + now("yyyyMMddHHmmss"); // add ss or ssSSSS if you wish
				
		// Find custom xml item id and inject data_file.xml		
		CustomXmlDataStoragePart customXmlDataStoragePart 
			= CustomXmlDataStoragePartSelector.getCustomXmlDataStoragePart(wordMLPackage);		
		if (customXmlDataStoragePart==null) {
			throw new WebApplicationException(
					new Docx4JException("Couldn't find CustomXmlDataStoragePart"), 
					Status.UNSUPPORTED_MEDIA_TYPE);
		}
		customXmlDataStoragePart.getData().setDocument(xis);
	
//		customXmlDataStoragePart.getData().setDocument(new ByteArrayInputStream(bytes));
		
		
//		java.lang.NullPointerException
//		org.docx4j.openpackaging.parts.XmlPart.setDocument(XmlPart.java:129)

		final SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
		
		// Process conditionals and repeats
		OpenDoPEHandler odh = new OpenDoPEHandler(wordMLPackage);
		odh.preprocess();

		OpenDoPEIntegrity odi = new OpenDoPEIntegrity();
		odi.process(wordMLPackage);
		
		if (log.isDebugEnabled()) {		
			try {
				File save_preprocessed = new File( System.getProperty("java.io.tmpdir") 
						+ "/" + filePrefix + "_INT.docx");
				saver.save(save_preprocessed);
				log.debug("Saved: " + save_preprocessed);
			} catch (Exception e) {
				log.error(e.getMessage(), e);
			}
		}
		
		// Apply the bindings
		// Convert hyperlinks, using this style
		BindingHandler.setHyperlinkStyle(hyperlinkStyleId);	
		
		AtomicInteger bookmarkId = odh.getNextBookmarkId();
		BindingHandler bh = new BindingHandler(wordMLPackage);
		bh.setStartingIdForNewBookmarks(bookmarkId);

		bh.applyBindings();
		
		if (log.isDebugEnabled()) {			
			try {
				File save_bound = new File( System.getProperty("java.io.tmpdir") 
						+ "/" + filePrefix + "_BOUND.docx"); 
				saver.save(save_bound);
				log.debug("Saved: " + save_bound);
			} catch (Exception e) {
				log.error(e.getMessage(), e);
			}
		}		
		
		
		// Update TOC .. before RemovalHandler!  (why? because of ToC SDT)
		if (processTocField) {
			
			log.debug("toc processing requested");
			
			// Use reflection, so this WAR can be built
			// by users who don't have the Enterprise jar
			try {
				Class<?> tocGenerator = Class.forName("org.docx4j.toc.TocGenerator");
				
				Constructor ctor = tocGenerator.getDeclaredConstructor(WordprocessingMLPackage.class);
				Object tocGeneratorObj = ctor.newInstance(wordMLPackage);				
				
				//Method method = documentBuilder.getMethod("merge", wmlPkgList.getClass());			
				Method[] methods = tocGenerator.getMethods(); 
				Method methodUpdateToc = null;
				Method methodSetDocumentServicesEndpoint = null;
				Method methodSetStartingIdForNewBookmarks = null;
				for (int j=0; j<methods.length; j++) {
//					System.out.println(methods[j].getName());
					if (methods[j].getName().equals("updateToc")
							&& methods[j].getParameterTypes().length==1) {
						methodUpdateToc = methods[j];
					} else if (methods[j].getName().equals("setDocumentServicesEndpoint")) {
						methodSetDocumentServicesEndpoint = methods[j];
					} else if (methods[j].getName().equals("setStartingIdForNewBookmarks")) {
						methodSetStartingIdForNewBookmarks = methods[j];
					}

				}			
				if (methodUpdateToc==null) {
					log.error("toc processing requested, but Enterprise jar not available");				
				} else {
				
					Document contentBackup = XmlUtils.deepCopy(wordMLPackage.getMainDocumentPart().getJaxbElement());
					try {
						if (documentServicesEndpoint!=null) {
							methodSetDocumentServicesEndpoint.invoke(tocGeneratorObj, documentServicesEndpoint);
						}
						methodSetStartingIdForNewBookmarks.invoke(tocGeneratorObj, bookmarkId);
						methodUpdateToc.invoke(tocGeneratorObj, !tocPageNumbers);
						
					} catch (Exception e1) {
						log.error(e1.getMessage(), e1);
						log.error("Omitting TOC; generate that in Word");
						wordMLPackage.getMainDocumentPart().setJaxbElement(contentBackup);
//						TocGenerator.updateToc(wordMLPackage, true);
					}
				}
				
			} catch (Exception e) {
				log.error("toc processing requested, but Enterprise jar not available");				
				log.error(e.getMessage(), e);
			}			
		}
		
		// Strip content controls: you MUST do this 
		// if you are processing hyperlinks
		RemovalHandler rh = new RemovalHandler();
		rh.removeSDTs(wordMLPackage, Quantifier.ALL);
		
		
		
		if (log.isDebugEnabled()) {
			try {
				File save = new File( System.getProperty("java.io.tmpdir") 
						+ "/" + filePrefix + "_STRIPPED.docx"); 
				saver.save(save);
				log.debug("Saved: " + save);
			} catch (Exception e) {
				log.error(e.getMessage(), e);
			}
		}
		
				
		if (format!=null) {
			
			if (format.equals("html") ) {		
		
				// Return html
				final AbstractHtmlExporter exporter = new HtmlExporterNG2(); 	
		    	final HtmlSettings htmlSettings = new HtmlSettings();
		    	htmlSettings.setImageDirPath(htmlImageDirPath);	    	
				htmlSettings.setImageTargetUri(htmlImageTargetUri); 
		
				ResponseBuilder builder = Response.ok(
				
					new StreamingOutput() {
						
						public void write(OutputStream output) throws IOException, WebApplicationException {
							
							javax.xml.transform.stream.StreamResult result 
								= new javax.xml.transform.stream.StreamResult(output);
							
							try {
								exporter.html(wordMLPackage, result, htmlSettings);
							} catch (Exception e) {
								throw new WebApplicationException(e,
										Status.INTERNAL_SERVER_ERROR);
							}
							
						}
					}
				);
				builder.type("text/html");
				return builder.build();
				
			} else if (format.equals("pdf") ) {	
				
				if (documentServicesEndpoint!=null) {
					
					final ByteArrayOutputStream tmpDocxFile = new ByteArrayOutputStream(); 
					try {
						Docx4J.save(wordMLPackage, tmpDocxFile, Docx4J.FLAG_SAVE_ZIP_FILE);
					} catch (Exception e) {
						throw new TocException("Error saving pkg as tmp file; " + e.getMessage(),e);
					}   
					
			    	
			    	// 
					final Converter c = new ConverterHttp(documentServicesEndpoint); 

					ResponseBuilder builder = Response.ok(
							
							new StreamingOutput() {				
								public void write(OutputStream output) throws IOException, WebApplicationException {					
							         try {
							 			c.convert(tmpDocxFile.toByteArray(), Format.DOCX, Format.PDF, output);
									} catch (Exception e) {
										throw new WebApplicationException(e);
									} 					
								}
							}
						);
		//						builder.header("Content-Disposition", "attachment; filename=output.pdf");
						builder.type("application/pdf");
						
						return builder.build();
					
				} else {
				
					final org.docx4j.convert.out.pdf.PdfConversion c 
						= new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(wordMLPackage);
		
					ResponseBuilder builder = Response.ok(
						
						new StreamingOutput() {				
							public void write(OutputStream output) throws IOException, WebApplicationException {					
						         try {
						 			c.output(output, new PdfSettings() );
								} catch (Docx4JException e) {
									throw new WebApplicationException(e);
								}							
							}
						}
					);
	//						builder.header("Content-Disposition", "attachment; filename=output.pdf");
					builder.type("application/pdf");
					
					return builder.build();
				}

			} else if (format.equals("docx") ) {		
				
				// fall through to the below
				
			} else {
				log.error("Unkown format: " + format);
				return Response.ok("<p>Unknown format " + format + " </p>",
						MediaType.APPLICATION_XHTML_XML_TYPE).build();				
				
			}
		}
		
		// return a docx
		ResponseBuilder builder = Response.ok(
				
			new StreamingOutput() {				
				public void write(OutputStream output) throws IOException, WebApplicationException {					
			         try {
						saver.save(output);
					} catch (Docx4JException e) {
						throw new WebApplicationException(e);
					}							
				}
			}
		);
		builder.header("Content-Disposition", "attachment; filename=" + filePrefix + ".docx");
		builder.type("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
		return builder.build();
         
	}
	
	/**
	 * header sample
	 * {
	 * 	Content-Type=[image/png], 
	 * 	Content-Disposition=[form-data; name="file"; filename="filename.extension"]
	 * }
	 **/
	private String getFileName(String name) {
 
		String finalFileName = name.trim().replaceAll("\"", "");
		
		if (finalFileName.lastIndexOf(".")==-1) {
			return finalFileName;
		} else {
			return finalFileName.substring(0, finalFileName.lastIndexOf(".") ); 
		}
     }

	public static String now(String dateFormat) {
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
		return sdf.format(cal.getTime());
	}
	
//    public static String getDefaultEncoding() {
//    	
////    	try {
////        	System.setProperty("file.encoding","UTF-8");
////        	Field charset = Charset.class.getDeclaredField("defaultCharset");
////        	charset.setAccessible(true);
////			charset.set(null,null);
////		} catch (Exception e) {
////			// TODO Auto-generated catch block
////			e.printStackTrace();
////		}  	
//    	
//    	byte [] byteArray = {'a'}; 
//    	InputStream inputStream = new ByteArrayInputStream(byteArray); 
//    	InputStreamReader reader = new InputStreamReader(inputStream); 
//    	return reader.getEncoding();     	
//    	
//    }
	
	
}
