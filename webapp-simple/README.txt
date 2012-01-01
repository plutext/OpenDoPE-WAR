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

<tomcat-users>
  <role rolename="manager"/>
  <user username="test" password="test" roles="manager"/>
  

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

Start tomcat

Then you can do something like:

    mvn clean tomcat:deploy