package main.utils;

import java.io.File;
import java.io.FileFilter;

public class CommonUtil {

	public File[] getFiles(String folderPath) {
		File folder = new File(folderPath);
		File[] listOfFiles = folder.listFiles(new FileFilter() {          
		    public boolean accept(File file) {
		        return file.isFile();
		    }
		});
		return listOfFiles;
	}
}
