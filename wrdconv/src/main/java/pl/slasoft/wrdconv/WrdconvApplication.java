package pl.slasoft.wrdconv;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.WebApplicationType;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;

import pl.slasoft.wrdconv.service.ServiceOne;

@SpringBootApplication
public class WrdconvApplication {

	public static void main(String[] args) {
		System.out.println("Word DOCs converted starting...");
		
        SpringApplication app = new SpringApplication(WrdconvApplication.class);
        app.setWebApplicationType(WebApplicationType.NONE); // Disable web server
        ApplicationContext context = app.run(args);
        
        ServiceOne srvOne = context.getBean(ServiceOne.class);
        srvOne.transformWordFiles();
        //srvOne.proc01();
	}

}
