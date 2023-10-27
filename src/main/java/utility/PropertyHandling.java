package utility;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import static java.lang.System.*;


public class PropertyHandling {
    Properties properties;
    public PropertyHandling(){
        try {
            // Path of the properties file.
            String filePath= (".//target//config.properties");

            // To read the file.
            FileInputStream file = new FileInputStream(filePath);

            //Initialising the properties.
            properties = new Properties();

            //Load the file in the properties class to retrieve the values
            properties.load(file);
        }
        catch (IOException e){
            throw new RuntimeException(e);
        }
    }
    public String getProperty(String key){
        String value = properties.getProperty(key);
        out.println(value);
        return value;
    }

    public static void main(String[] args) {

        String browser =new PropertyHandling().getProperty("browser");
        String url = new PropertyHandling().getProperty("orangeHrmUrl");

        out.println("The browser to open " + url + " is : " + browser);
    }
}
