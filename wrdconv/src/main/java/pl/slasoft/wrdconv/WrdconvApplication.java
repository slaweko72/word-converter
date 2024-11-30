package pl.slasoft.wrdconv;

import java.util.Arrays;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.WebApplicationType;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class WrdconvApplication {

	public static void main(String[] args) {
		System.out.println("Word converted starting...");
		System.out.println("args: " + Arrays.toString(args));
		
        SpringApplication app = new SpringApplication(WrdconvApplication.class);
        app.setWebApplicationType(WebApplicationType.NONE); // Disable web server
        app.run(args);        
	}

}
