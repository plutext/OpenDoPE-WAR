package org.opendope.webapp;

import java.util.HashSet;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;


import javax.servlet.ServletContext;
import javax.ws.rs.core.Application;
import javax.ws.rs.core.Context;

import org.glassfish.jersey.filter.LoggingFilter;
import org.glassfish.jersey.media.multipart.MultiPartFeature;


/**
 * Portable JAX-RS application.
 */
public class JAXRS2Application extends Application {
	
	private static final Logger jul = Logger.getLogger(JAXRS2Application.class
			.getName());
	
	static {
		
		//Logger.getLogger("com.sun.jersey").setLevel(Level.FINEST);
		Logger.getLogger("org.glassfish.jersey").setLevel(Level.FINEST);
		Logger.getLogger("org.plutext").setLevel(Level.FINEST);
	}
		
    @Context 
    private ServletContext servletConfig; // can't be static
    
//    

//    // In order that implementations can add their own params, without having to edit this class
//    public static ServletContext getServletConfig() {
//		return servletConfigStatic;
//	}
//    private static ServletContext servletConfigStatic; // (!)



    
    private static String contextPath;
    public static String getContextPath() {
		return contextPath;
	}



	
	
    @Override
    public Set<Class<?>> getClasses() {
    	
    	// Pattern @author Arul Dhesiaseelan (aruld@acm.org)
    	
        final Set<Class<?>> classes = new HashSet<Class<?>>();
        // register resources and features
        classes.add(MultiPartFeature.class);
        classes.add(SubmitBoth.class);
        

        
        
        classes.add(LoggingFilter.class);
        
        
        return classes;
    }
}
