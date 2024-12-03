package pl.slasoft.wrdconv.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
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
            logger.info("Number .doc* files found: " + docFiles.size());

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
            
            // Placeholder-value pairs to replace
            HashMap<String, String> placeholders = new HashMap<>();
            placeholders.put("${DATE}", config.getDate());
            placeholders.put("${MONTH}", config.getMonth());
            
            for (Path docFile : docFiles) {
            	replacePlaceholders(docFile.toString(), outputFolder.toString()+"\\"+docFile.getFileName(), placeholders);
            }
            
            //docFiles.forEach(dc -> replacePlaceholders(dc.toString(), null, null));            

        } catch (Exception e) {
        	logger.error("transformWordFiles error: " + e.getMessage());
        }		
	}
	
    // Method to replace placeholders in a Word document
    private void replacePlaceholders(String inputFilePath, String outputFilePath, HashMap<String, String> placeholders) 
            throws IOException {
        // Load the Word document
        XWPFDocument document = new XWPFDocument(new java.io.FileInputStream(inputFilePath));

        // Iterate over all paragraphs in the document
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    System.out.println("--> " + text);
                    if (text != null) {
                        for (String placeholder : placeholders.keySet()) {
                            if (text.contains(placeholder)) {
                                // Replace the placeholder with the actual value
                                text = text.replace(placeholder, placeholders.get(placeholder));
                                run.setText(text, 0); // Replace the text in the run
                            }
                        }
                    }
                }
            }
        }

        // Save the modified document to a new location
        try (FileOutputStream out = new FileOutputStream(new File(outputFilePath))) {
            document.write(out);
            System.out.println("Modified document saved to: " + outputFilePath);
        }

        // Close the document
        document.close();
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
