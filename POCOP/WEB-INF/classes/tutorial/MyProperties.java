package tutorial;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class MyProperties {
	
	public String driver=null;
	public String url=null;
	public String username=null;
	public String password=null;
	public String out_file=null;
	public String empty_file=null;
	public String myInput=null;
	public MyProperties() throws IOException {
	InputStream path = MyProperties.class.getResourceAsStream("filename.properties");
    Properties prop = new Properties();
    prop.load(path);
    driver=prop.getProperty("driver");
	url=prop.getProperty("url");
	username=prop.getProperty("username");
	password=prop.getProperty("password");
	out_file=prop.getProperty("output_file");
	empty_file=prop.getProperty("empty_file");
	myInput=prop.getProperty("myInput");
	
	}
	
}
