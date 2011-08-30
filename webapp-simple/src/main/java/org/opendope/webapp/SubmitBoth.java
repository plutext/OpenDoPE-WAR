/*
 *  Copyright 2010-11, Plutext Pty Ltd.
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
import java.util.List;
import java.util.Map;

import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.WebApplicationException;
import javax.ws.rs.core.MultivaluedMap;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;
import javax.ws.rs.core.Response.Status;
import javax.ws.rs.core.StreamingOutput;
import javax.xml.bind.JAXBContext;

import org.apache.log4j.Logger;
import org.docx4j.convert.out.html.AbstractHtmlExporter;
import org.docx4j.convert.out.html.AbstractHtmlExporter.HtmlSettings;
import org.docx4j.convert.out.html.HtmlExporterNG2;
import org.docx4j.model.datastorage.BindingHandler;
import org.docx4j.model.datastorage.OpenDoPEHandler;
import org.docx4j.model.datastorage.RemovalHandler;
import org.docx4j.model.datastorage.RemovalHandler.Quantifier;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.LoadFromZipNG;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.jboss.resteasy.plugins.providers.multipart.InputPart;
import org.jboss.resteasy.plugins.providers.multipart.MultipartFormDataInput;
import org.opendope.CustomXmlUtils;

@Path("/both")  // must match form action
public class SubmitBoth {

	final static String KEY_XML = "xmlfile";
	final static String KEY_DOCX = "docx";

	public static JAXBContext context = org.docx4j.jaxb.Context.jc; 
	
	private static Logger log = Logger.getLogger(SubmitBoth.class);		
	
	static {
		// Convert hyperlinks, using this style
		BindingHandler.setHyperlinkStyle("Hyperlink");
	}
		
	/**
	 * Display a form in which user can provide OpenDoPE docx template,
	 * and xml data file, for processing.  Useful for testing.
	 * The URL will be something like http://localhost:8080/OpenDoPE-simple/service/both
	 */
	@GET
	@Produces("text/html")
	public String uploadForm() {
		return "<html><head><title>OpenDoPE injection</title></head>"
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

	@POST
	@Consumes("multipart/form-data")
	@Produces( {"application/vnd.openxmlformats-officedocument.wordprocessingml.document" , 
				"text/html"})
	public Response processForm(MultipartFormDataInput input, @QueryParam("format") String format ) throws Docx4JException, IOException {
		
		log.info("requested format: " + format);
				
		final WordprocessingMLPackage wordMLPackage;
		String docxname = "";

		InputStream xis = null;
		String dataname = "";
		
		Map<String, List<InputPart>> map = input.getFormDataMap();

		WordprocessingMLPackage tmpPkg=null;
		for (String s : map.keySet()) {

			if (s.equals(KEY_XML)) {
				InputPart part = map.get(s).get(0);
				
				MultivaluedMap<String, String> header = part.getHeaders();
				dataname = getFileName(header);
				
				try {
					xis = part.getBody(InputStream.class,null);
				} catch (IOException e) {
					throw new WebApplicationException(
							new Docx4JException("Error accepting xml data stream"), 
							Status.BAD_REQUEST);
				}
			}
			
			if (s.equals(KEY_DOCX)) {
				InputPart part = map.get(s).get(0);

				MultivaluedMap<String, String> header = part.getHeaders();
				docxname = getFileName(header);
				
				LoadFromZipNG loader = new LoadFromZipNG();
				try {
					tmpPkg = (WordprocessingMLPackage)loader.get(
							part.getBody(InputStream.class,null) );
				} catch (Exception e) {
					throw new WebApplicationException(
							new Docx4JException("Error reading docx file (is this a valid docx?)"), 
							Status.BAD_REQUEST);
				}
			}
			
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
		String filePrefix = docxname + "_" + dataname + "_" + now("yyyyMMddHHmm"); // add ss or ssSSSS if you wish
				
		// Find custom xml item id
		String itemId = CustomXmlUtils.getCustomXmlItemId(wordMLPackage).toLowerCase();
		System.out.println("Looking for item id: " + itemId);
		
		// Inject data_file.xml		
		CustomXmlDataStoragePart customXmlDataStoragePart 
			= wordMLPackage.getCustomXmlDataStorageParts().get(itemId);
		if (customXmlDataStoragePart==null) {
			throw new WebApplicationException(
					new Docx4JException("Couldn't find CustomXmlDataStoragePart"), 
					Status.UNSUPPORTED_MEDIA_TYPE);
		}
		customXmlDataStoragePart.getData().setDocument(xis);

		final SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
		
		// Process conditionals and repeats
		OpenDoPEHandler odh = new OpenDoPEHandler(wordMLPackage);
		odh.preprocess();
		if (log.isDebugEnabled()) {			
			File save_preprocessed = new File( System.getProperty("java.io.tmpdir") 
					+ "/" + filePrefix + "_INT.docx");
			saver.save(save_preprocessed);
			log.debug("Saved: " + save_preprocessed);
		}
		
		// Apply the bindings
		BindingHandler.applyBindings(wordMLPackage);
		if (log.isDebugEnabled()) {			
			File save_bound = new File( System.getProperty("java.io.tmpdir") 
					+ "/" + filePrefix + "_BOUND.docx"); 
			saver.save(save_bound);
			log.debug("Saved: " + save_bound);
			
//			System.out.println(
//			XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true)
//			);			
		}		
		
		// Strip content controls: you MUST do this 
		// if you are processing hyperlinks
		RemovalHandler rh = new RemovalHandler();
		rh.removeSDTs(wordMLPackage, Quantifier.ALL);
		if (log.isDebugEnabled()) {			
			File save = new File( System.getProperty("java.io.tmpdir") 
					+ "/" + filePrefix + "_STRIPPED.docx"); 
			saver.save(save);
			log.debug("Saved: " + save);
		}		
				
		if (format!=null && format.equals("html") ) {		
			// Return html
			final AbstractHtmlExporter exporter = new HtmlExporterNG2(); 	
	    	final HtmlSettings htmlSettings = new HtmlSettings();
	    		// NB: as it stands, this quick html demo isn't intended to handle images
	    		// and hasn't been tested with them
	
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
	private String getFileName(MultivaluedMap<String, String> header) {
 
		String[] contentDisposition = header.getFirst("Content-Disposition").split(";");
		 
		for (String filename : contentDisposition) {
			if ((filename.trim().startsWith("filename"))) {
 
				String[] name = filename.split("=");
 
				String finalFileName = name[1].trim().replaceAll("\"", "");
				
				if (finalFileName.lastIndexOf(".")==-1) {
					return finalFileName;
				} else {
					return finalFileName.substring(0, finalFileName.lastIndexOf(".") ); 
				}
				
			}
		}
		return "unknown";
     }

	public static String now(String dateFormat) {
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
		return sdf.format(cal.getTime());
	}
	
}
