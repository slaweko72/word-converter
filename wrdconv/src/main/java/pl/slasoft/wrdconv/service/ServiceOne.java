package pl.slasoft.wrdconv.service;

import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import pl.slasoft.wrdconv.config.CustomConfig;

@Service
public class ServiceOne {

	private static final Logger logger = LoggerFactory.getLogger(ServiceOne.class);

	private final CustomConfig config;

	public ServiceOne(CustomConfig config) {
		super();
		this.config = config;
	}	
	
	public String getDate() {
		return config.getDate() + " - " + config.getMonth();
	}
	
	public void proc01() {
		SimpleDateFormat dtf = new SimpleDateFormat("yyyyMMdd_HHmmss");
		Date dt = new Date();
		String folderName = "_OUTPUT_" + dtf.format(dt);
		logger.info("--------------------> " + folderName);
	}
		
	public void transformWordFiles() {
		logger.info("transformWordFiles started");

        try {
            // Get the current folder
            Path currentFolder = Paths.get(System.getProperty("user.dir"));
    		logger.info("currentFolder: " + currentFolder);

            // Find .doc* files in the current folder
            List<Path> docFiles = findDocFiles(currentFolder);            
            logger.info(".doc* files found: " + docFiles.toString());

//            // Print the list of .doc files
//            if (docFiles.isEmpty()) {
//            	logger.info("No .doc* files found in the current folder.");
//            } else {
//            	logger.info("Found the following .doc* files:");
//                docFiles.forEach(dc -> logger.info(dc.toString()));
//            }
            
            // Create output folder
            File outputFolder = createOutputFolder(currentFolder);
            logger.info("Output folder created: " + outputFolder.toString());

        } catch (Exception e) {
        	logger.error("transformWordFiles error: " + e.getMessage());
        }		
	}
	
	private File createOutputFolder(Path currentFolder) throws Exception {
		String folderName = getFolderName();		
		File newDirectory = new File(currentFolder.toFile(), folderName); 
		if (!newDirectory.mkdir()) {
			throw new Exception("Can't create directory: " + folderName);
		}		
		return newDirectory;
	}
	
    private List<Path> findDocFiles(Path folder) throws IOException {
        List<Path> docFiles = new ArrayList<>();

        // Use DirectoryStream to filter files by extension
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(folder, "*.doc*")) {
            for (Path entry : stream) {
                docFiles.add(entry);
            }
        }

        return docFiles;
    }
    
	public String getFolderName() {
		SimpleDateFormat dtf = new SimpleDateFormat("yyyyMMdd_HHmmss");
		Date dt = new Date();
		return "_OUTPUT_" + dtf.format(dt);
	}
    
		
}
