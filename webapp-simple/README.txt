This is a webapp, so it needs to be installed in a servlet container.

Once you've installed it in (for example) tomcat, in a browser, go to:

http://localhost:8080/OpenDoPE-simple/service/both

You should see a simple form.  There is a sample xml file and docx you can try out; get these from: 

https://github.com/plutext/OpenDoPE-WAR/tree/master/webapp-simple/samples/invoice

Maven
-----

You can use maven to automatically deploy to tomcat.

First install Tomcat 6. 
			        	  			        	  
Configure username and password in Maven settings.xml, and tomcat-users.xml

<settings xmlns="http://maven.apache.org/POM/4.0.0"
   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
   xsi:schemaLocation="http://maven.apache.org/POM/4.0.0
                      http://maven.apache.org/xsd/settings-1.0.0.xsd">
  <servers>
    <server>
      <id>tomcat-localhost</id>
      <username>test</username>
      <password>test</password>
    </server>
  </servers>
</settings>

Tomcat < 6.0.30
---------------

<tomcat-users>
  <role rolename="manager"/>
  <user username="test" password="test" roles="manager"/>
</tomcat-users>

  
Tomcat 6.0.30 onwards
---------------

<tomcat-users>
  <role rolename="manager-gui"/>
  <role rolename="manager-script"/>
  <user username="test" password="test" roles="manager-gui,manager-script"/>  
</tomcat-users>



Start tomcat

Then you can do something like:

    mvn clean tomcat6:deploy

    mvn tomcat6:undeploy
    
troubleshooting maven
---------------------

1. Make sure you are using tomcat6:deploy (note the '6'), otherwise old tomcat plugin will be used
2. For Tomcat *6.0.30 onwards*, you need
   (i) role manager-script (manager-gui also useful)
   (ii)<url>http://localhost:8080/manager/script</url>  (note 'script')

troubleshooting tomcat 
---------------------

java.lang.OutOfMemoryError: PermGen space (eg when ticking html output box) 

easiest if you have tomcat installed as a service.

Run cmd as administrator

In tomcat bin dir:

   service install tomcat6
   
then run tomcat6w as administrator

On the Java tab, I added at the beginning of Java options:

   -XX:MaxPermSize=512M
   
and set 'maximum memory pool' to 1024MB


    
web.xml
-------

configuration via web.xml.

You'll see 3 init params there:

        <init-param>
            <param-name>HtmlImageTargetUri</param-name>
            <param-value>/OpenDoPE-simple/</param-value>
        </init-param>
        <init-param>
            <param-name>HtmlImageDirPath</param-name>
            <param-value>/apache-tomcat-7.0.20/webapps/OpenDoPE-simple/</param-value>
        </init-param>
        <init-param>
            <param-name>HyperlinkStyleId</param-name>
            <param-value>Hyperlink</param-value>
        </init-param>

The first two are for images in html output.  They allow you to say where docx4j should write the image, 
and the corresponding path which will appear in the html page. With this, you should now see images in the html output.

Note that if your webapp is named OpenDoPE-simple-1.0.0, you'll want to make the init params match that (ie add "-1.0.0")

The other parameter allows you to specify the name of the style which is to be used for hyperlinks. docx4j gets the 
actual definition from styles known to it, so you probably won't want to change this (ie unless you change those 
style definitions in docx4j).    

