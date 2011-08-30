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
package org.opendope;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

public class CustomXmlUtils {
	
	/**
	 * We need the item id of the custom xml part.  
	 * 
	 * There are several strategies we could use to find it,
	 * including searching the docx for a bind element, but
	 * here, we simply look in the xpaths part. 
	 * 
	 * @param wordMLPackage
	 * @return
	 */
	public static String getCustomXmlItemId(WordprocessingMLPackage wordMLPackage) throws Docx4JException {
		
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();			
		if (wordMLPackage.getMainDocumentPart().getXPathsPart()==null) {
			throw new Docx4JException("OpenDoPE XPaths part missing");
		} 
	
		org.opendope.xpaths.Xpaths xPaths = wordMLPackage.getMainDocumentPart().getXPathsPart().getJaxbElement();
		
		return xPaths.getXpath().get(0).getDataBinding().getStoreItemID();
		
	}
	

}
