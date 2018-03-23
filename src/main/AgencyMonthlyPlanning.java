package main;

import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;

import utils.XLSXReaderWriter;

public class AgencyMonthlyPlanning {
	private static XLSXReaderWriter xlsxRW = new XLSXReaderWriter();
	private static File[] getFiles(String folderPath) {
		File folder = new File(folderPath);
		File[] listOfFiles = folder.listFiles(new FileFilter() {          
		    public boolean accept(File file) {
		        return file.isFile();
		    }
		});
		return listOfFiles;
	}
	/**
	 * 
	 * @param rowidx
	 * @param st: pathname string
	 * @param teamName: data from source will be filtered by team name
	 * @param exportedFile: destination file where the output data will be saved
	 * @return row index: the index of the last row in the excel output file
	 * @throws IOException
	 */
	/*
	private static int readFilterWriteData(String st, String teamName, String exportedFile, int rowidx) throws IOException {
		ArrayList<Row> rows = xlsxRW.read(st, "Individuals_Detail Plan", 10);
		if (rowidx == 0) {
			Row headerRow = rows.get(0);
			ArrayList<Row> header = new ArrayList<Row>();
			header.add(headerRow);
			rowidx = xlsxRW.write(exportedFile, "Indi", rowidx, header);
		} 
		ArrayList<Row> newData = (xlsxRW.filter(rows, teamName, 2)); // filter data by team name in the column 2 nd
		rowidx = xlsxRW.write(exportedFile, "Indi", rowidx, newData);
		return rowidx;
	}
	*/
	/**
	 * 
	 * @param rowidx
	 * @param st: pathname string
	 * @param teamName: data from source will be filtered by team name
	 * @param exportedFile: destination file where the output data will be saved
	 * @return row index: the index of the last row in the excel output file
	 * @throws IOException
	 */
	private static int filterWriteData(ArrayList<Row> rows, String teamName, String exportedFile, int rowidx, String sheetname, int filteredColIndex) throws IOException {
		if (rowidx == 0) {
			Row headerRow = rows.get(0);
			ArrayList<Row> header = new ArrayList<Row>();
			header.add(headerRow);
			rowidx = xlsxRW.write(exportedFile, sheetname, rowidx, header);
		} 
		ArrayList<Row> newData = (xlsxRW.filter(rows, teamName, filteredColIndex)); // filter data by team name in the column 2 nd
		rowidx = xlsxRW.write(exportedFile, sheetname, rowidx, newData);
		return rowidx;
	}

	public static void main(String[] args) {
		long startTime = System.currentTimeMillis();
		String exportedFile = "d:\\workspace_excel\\PHUCANH\\Agency Monthly Sales Plan 201803\\output\\out.xlsx";
		try {
			// get all files from the folder
			File[] files = getFiles("d:\\workspace_excel\\PHUCANH\\Agency Monthly Sales Plan 201803\\");
			int rowidxIndiSheet = 0;
			int rowidxUnitSheet = 0;
			int rowidxSessSheet = 0;
			
//			ArrayList<Row> activities = xlsxRW.read("d:\\workspace_excel\\PHUCANH\\Agency Monthly Sales Plan 201803\\Agency Monthly Planning 201803 _ BÀ TRIỆU.xlsm", "Activity plan", 11, 2);
//			rowidxIndiSheet = filterWriteData(activities, "Sessions", exportedFile, rowidxIndiSheet, "Indi");
//			rowidxIndiSheet = filterWriteData(activities, "Attendees", exportedFile, rowidxIndiSheet, "Indi");
//			
//			if(false)
			for (int i=0; i<files.length; i++) {
				File f = files[i];
				if (!f.getName().startsWith("~")) {
					String st = f.getPath();
					String teamName = st.split(" _ ")[1].split(".xls")[0];
					System.out.println(String.format("%d. Reading for: %s", i, teamName));
					
					/*ArrayList<Row> indiRows = xlsxRW.read(st, "Individuals_Detail Plan", 10);
					ArrayList<Row> leaderRows = xlsxRW.read(st, "Leaders_Detail Plan", 13);
					rowidxIndiSheet = filterWriteData(indiRows, teamName, exportedFile, rowidxIndiSheet, "Indi");
					rowidxUnitSheet = filterWriteData(leaderRows, teamName, exportedFile, rowidxUnitSheet, "Unit");*/
					ArrayList<Row> activities = xlsxRW.read(f, "Activity plan", 11, 3);
					rowidxSessSheet = filterWriteData(activities, "Sessions", exportedFile, rowidxSessSheet, "Session", 1);
					
					
//					Map<String, Object> dataDic = xlsxRW.read(f, new String[] {"Individuals_Detail Plan", "Leaders_Detail Plan"}, new int[] {10, 13});
//					Set<String> keySet = dataDic.keySet();
//					for (String k: keySet) {
//						ArrayList<Row> rows = (ArrayList<Row>)dataDic.get(k);
//						if (k.equals("Individuals_Detail Plan")) {
//							rowidxIndiSheet = filterWriteData(rows, teamName, exportedFile, rowidxIndiSheet, "Indi", 2);
//						} else if (k.equals("Leaders_Detail Plan")) {
//							rowidxUnitSheet = filterWriteData(rows, teamName, exportedFile, rowidxUnitSheet, "Unit", 2);
//						}
//					}
				}
			}

			long stopTime = System.currentTimeMillis();
			long elapsedTime = stopTime - startTime;
			// Converting Milliseconds to Minutes and Seconds
			long minutes = TimeUnit.MILLISECONDS.toMinutes(elapsedTime);
			long seconds = TimeUnit.MILLISECONDS.toSeconds(elapsedTime);
			System.out.println(String.format("Done in %d minute(s) %d seconds.", minutes, seconds));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
