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

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.annotation.PostConstruct;
import javax.servlet.ServletConfig;
import javax.servlet.ServletContext;
import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.WebApplicationException;
import javax.ws.rs.core.Context;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;
import javax.ws.rs.core.Response.Status;
import javax.ws.rs.core.StreamingOutput;
import javax.xml.bind.JAXBContext;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporter;
import org.docx4j.convert.out.html.AbstractHtmlExporter;
import org.docx4j.convert.out.html.AbstractHtmlExporter.HtmlSettings;
import org.docx4j.convert.out.html.HtmlExporterNG2;
import org.docx4j.model.datastorage.BindingHandler;
import org.docx4j.model.datastorage.OpenDoPEHandler;
import org.docx4j.model.datastorage.OpenDoPEIntegrity;
import org.docx4j.model.datastorage.RemovalHandler;
import org.docx4j.model.datastorage.RemovalHandler.Quantifier;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.LoadFromZipNG;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.docx4j.wml.RFonts;
import org.opendope.CustomXmlUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sun.jersey.core.header.FormDataContentDisposition;
import com.sun.jersey.multipart.FormDataParam;

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
    
    @PostConstruct
    public void readInitParams() {
    	
    	// This works with and without class extends Application
       
    	log.info( servletConfig.getInitParameter("HyperlinkStyleId") );
    	log.info( servletConfig.getInitParameter("HtmlImageTargetUri") );
    	log.info( servletConfig.getInitParameter("HtmlImageDirPath") );
    	
        hyperlinkStyleId = servletConfig.getInitParameter("HyperlinkStyleId");
        htmlImageTargetUri = servletConfig.getInitParameter("HtmlImageTargetUri");
        htmlImageDirPath = servletConfig.getInitParameter("HtmlImageDirPath");
        
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
			XHTMLImporter.setCssWhiteList(cssWhiteList);
		} catch (IOException e) {
			log.warn("CSS-WhiteList not found", e);
		}
		
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
				
				+ "<script type=\"text/javascript\">"
				+ "function OnSubmitForm()"
				+ "{"
				+ "  if(document.uploadForm.format.checked == true) {"
				+ "    document.uploadForm.action =\"both?format=html\";"
				+ "  } else  {"
				+ "    document.uploadForm.action =\"both\";"
				+ "  }"
				+ "  return true;"
				+ "}"
				+ "</script>"	
				
				+ "<form name=\"uploadForm\" action=\"both\" " 
				+		" onsubmit=\"return OnSubmitForm();\" "
				+ 		" method=\"post\" enctype=\"multipart/form-data\">"
				+ "<label for=\"" + KEY_XML + "\">XML data file:</label> "
				+ "<input type=\"file\" name=\"" + KEY_XML + "\" id=\"" + KEY_XML + "\" /> "
				+ "<br/>"
				+ "<label for=\"" + KEY_DOCX + "\">docx template:</label> "
				+ "<input type=\"file\" name=\"" + KEY_DOCX + "\" id=\"" + KEY_DOCX + "\" /> "
				+ "<br/>"
				+ "<input type=\"checkbox\" name=\"format\" value=\"HTML\" checked=\"true\">as HTML</input>"
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
	@Consumes("multipart/form-data")
	@Produces( {"application/vnd.openxmlformats-officedocument.wordprocessingml.document" , 
				"text/html"})
	public Response processForm(
//			@Context ServletContext servletContext, 
//			@Context  HttpServletRequest servletRequest,
			@FormDataParam("xmlfile") InputStream xis,
			@FormDataParam("xmlfile") FormDataContentDisposition xmlFileDetail,
			@FormDataParam("docx") InputStream docxInputStream,
			@FormDataParam("docx") FormDataContentDisposition docxFileDetail,
			@QueryParam("format") String format
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
		
		// XHTML Import: no need to add a mapping, if docx defaults are to be used
		//addFontMapping(String cssFontFamily, RFonts rFonts)
		
		
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
				
		// Find custom xml item id
		String itemId = CustomXmlUtils.getCustomXmlItemId(wordMLPackage).toLowerCase();
		System.out.println("Looking for item id: " + itemId);
		
		// Inject data_file.xml		
		CustomXmlDataStoragePart customXmlDataStoragePart 
			= (CustomXmlDataStoragePart)wordMLPackage.getCustomXmlDataStorageParts().get(itemId);
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
		BindingHandler.applyBindings(wordMLPackage);
		if (log.isDebugEnabled()) {			
			try {
				File save_bound = new File( System.getProperty("java.io.tmpdir") 
						+ "/" + filePrefix + "_BOUND.docx"); 
				saver.save(save_bound);
				log.debug("Saved: " + save_bound);
			} catch (Exception e) {
				log.error(e.getMessage(), e);
			}
			
//			System.out.println(
//			XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true)
//			);			
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
				
		if (format!=null && format.equals("html") ) {		
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
